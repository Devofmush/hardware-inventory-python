import wmi
import json
import os as os_module
import shutil
from datetime import datetime
import win32com.client
import subprocess

c = wmi.WMI()
system_info = c.Win32_ComputerSystem()[0]
computer_name = system_info.Name
domain_name = system_info.Domain if system_info.Domain else "Workgroup"
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

filename = f"{computer_name}_{domain_name}_{timestamp}.json"
wmi_info = {}

def safe_wmi_object_to_dict(wmi_object):
    result = {}
    for prop in wmi_object.properties:
        try:
            result[prop] = getattr(wmi_object, prop, None)
        except Exception as e:
            result[prop] = f"Error retrieving property: {str(e)}"
    return result

#CPU details
cpu_info = []
for cpu in c.Win32_Processor():
    cpu_info.append(safe_wmi_object_to_dict(cpu))
wmi_info['CPU'] = cpu_info

#motherboard details
motherboard_info = []
for board in c.Win32_BaseBoard():
    motherboard_info.append(safe_wmi_object_to_dict(board))
wmi_info['Motherboard'] = motherboard_info

#memory module details
memory_info = []
for mem in c.Win32_PhysicalMemory():
    memory_info.append(safe_wmi_object_to_dict(mem))
wmi_info['MemoryModules'] = memory_info

#printers
printer_info = []
for printer in c.Win32_Printer():
    printer_info.append({
        'DeviceID': printer.DeviceID,
        'DriverName': printer.DriverName,
        'Local': printer.Local,
        'Network': printer.Network,
        'PortName': printer.PortName,
        'PrinterStatus': 'Online' if printer.PrinterStatus == 3 else 'Offline'
    })
wmi_info['Printers'] = printer_info

#imaging devices from WIA
wia_service = win32com.client.Dispatch("WIA.DeviceManager")
wia_devices_info = []

for device in wia_service.DeviceInfos:
    wia_devices_info.append({
        'DeviceID': device.DeviceID,
        'Type': device.Type,
        'Name': device.Properties('Name').Value,
        'Description': device.Properties('Description').Value
    })

wmi_info['WIADevices'] = wia_devices_info if wia_devices_info else "No WIA devices found"

#DVD/CD-ROM
dvd_info = []
for dvd in c.Win32_CDROMDrive():
    dvd_info.append(safe_wmi_object_to_dict(dvd))
wmi_info['DVD/CD-ROM'] = dvd_info

#disk drives
disk_info = []
for disk in c.Win32_DiskDrive():
    disk_info.append(safe_wmi_object_to_dict(disk))
wmi_info['Disks'] = disk_info

#OS details
os_info = []
for os in c.Win32_OperatingSystem():
    os_info.append(safe_wmi_object_to_dict(os))
wmi_info['OperatingSystem'] = os_info

#BIOS details
bios_info = []
for bios in c.Win32_BIOS():
    bios_info.append(safe_wmi_object_to_dict(bios))
wmi_info['BIOS'] = bios_info

#network adapter details
def get_network_info_ipconfig():
    try:
        #decode using 'cp850' for Windows encoding - dont have a clue why
        ipconfig_output = subprocess.check_output('ipconfig /all', shell=True).decode('cp850')
        
        ipconfig_lines = ipconfig_output.splitlines()
        
        network_adapters = []
        current_adapter = {}
        
        for line in ipconfig_lines:
            if line.startswith('Ethernet adapter') or line.startswith('Wireless LAN adapter'):
                if current_adapter:
                    network_adapters.append(current_adapter)
                current_adapter = {'Name': line.strip(':')}
            elif 'Physical Address' in line:
                current_adapter['MACAddress'] = line.split(':')[-1].strip()
            elif 'IPv4 Address' in line or 'IPv6 Address' in line:
                ip_type = 'IPv4' if 'IPv4' in line else 'IPv6'
                current_adapter[ip_type] = line.split(':')[-1].strip().replace('(Preferred)', '')
            elif 'Description' in line:
                current_adapter['Description'] = line.split(':')[-1].strip()
            elif 'Default Gateway' in line:
                current_adapter['DefaultGateway'] = line.split(':')[-1].strip()

        if current_adapter:
            network_adapters.append(current_adapter)
        
        return network_adapters
    except Exception as e:
        return f"Error retrieving network information: {str(e)}"
    
wmi_info['NetworkAdapters'] = get_network_info_ipconfig()

#Windows SID
try:
    sid_output = subprocess.check_output('wmic csproduct get UUID', shell=True).decode().strip()
    sid_lines = sid_output.split('\n')
    
    if len(sid_lines) > 1:
        sid = sid_lines[1].strip()
    else:
        sid = "SID not found"
    
    wmi_info['WindowsSID'] = sid
except Exception as e:
    wmi_info['WindowsSID'] = f"Error retrieving SID: {e}"

#user profiles
def fetch_user_profiles():
    users_directory = os_module.path.expandvars(r'C:\Users')
    user_profiles = []
    try:
        for user_profile in os_module.listdir(users_directory):
            if os_module.path.isdir(os_module.path.join(users_directory, user_profile)):
                user_profiles.append(user_profile)
    except Exception as e:
        user_profiles.append(f"Error retrieving profiles: {e}")
    return user_profiles

wmi_info['LoggedInUsersHistory'] = fetch_user_profiles()

#JSON file
with open(filename, 'w') as f:
    json.dump(wmi_info, f, indent=4)

print(f"WMI info has been saved to '{filename}'")

# --- File Copy ---

PATH_ADDR = r'//*path*'

try:
    destination_path = PATH_ADDR

    shutil.copy(filename, destination_path)
    print(f"File copied to {destination_path}")

    print(f"Attempting to delete the original file: '{filename}' from the current directory.")

    if os_module.path.exists(filename):
        os_module.remove(filename)
        print(f"File '{filename}' deleted from the current directory.")
    else:
        print(f"File '{filename}' not found in the current directory.")
except Exception as e:
    print(f"An error occurred during file copy or deletion: {e}")
