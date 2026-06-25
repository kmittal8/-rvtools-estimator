#Imports the Pandas library, which is essential for data manipulation
import pandas as pd

#Defines a function named process_excel that takes a file path as input and performs various calculations on the Excel data.
def process_excel(file_path):
  df_vHost = pd.read_excel(file_path, sheet_name='vHost')
  df_vInfo = pd.read_excel(file_path, sheet_name='vInfo')
  #Reads two sheets named 'vHost' and 'vInfo' from the specified Excel file into Pandas DataFrames named df_vHost and df_vInfo respectively

  # Filter vHost DataFrame to exclude rows with "VMware Virtual Processor". This is to exclude the Dummy HCX Host from calculation. 
  vHost_filtered = df_vHost[df_vHost['CPU Model'] != 'VMware Virtual Processor']

  # CPU calculations
  total_cores = vHost_filtered['# Cores'].sum()
  avg_cpu_usage = round(vHost_filtered['CPU usage %'].mean(), 2)
  cores_based_on_avg_cpu = round((avg_cpu_usage / 100) * total_cores, 2)

  # Memory calculations
  total_memory_bytes = vHost_filtered['# Memory'].sum()
  total_memory_gb = round(total_memory_bytes / 1024, 2)  # Divide by 1024 once
  avg_mem_usage = round(vHost_filtered['Memory usage %'].mean(), 2)
  mem_based_on_avg_usage = round((avg_mem_usage / 100) * total_memory_gb, 2)

  # Count powered-on VMs
  powered_on_vms = df_vInfo[df_vInfo['Powerstate'] == 'poweredOn']
  count_powered_on_vms = len(powered_on_vms)

  # Count powered-on Windows VMs from powered-on VMs. Why- OS according to the configuration file, coz VMware Tools are NOT always Updated.
  
  powered_on_windows_vms = powered_on_vms[powered_on_vms['OS according to the configuration file'].str.contains('Windows', case=False)]
  count_powered_on_windows_vms = len(powered_on_windows_vms)

  # Calculate total provisioned disk size for powered-on VMs. Remember: Provisioned MiB includes: Total disk capacity MiB. 
  total_provisioned_mib = powered_on_vms['Provisioned MiB'].sum()
  disk_intb=round(total_provisioned_mib / 1024 /1024, 2)

  print(f"File: {file_path}")
  print("Total Cores:", total_cores)
  print("Average CPU Usage % :", avg_cpu_usage)
  print(f"{avg_cpu_usage}% of total cores is: {cores_based_on_avg_cpu} Cores")
  print("Total Memory:", total_memory_gb, "GB")
  print("Average Memory Usage:", avg_mem_usage)
  print(f"{avg_mem_usage}% of total memory is: {mem_based_on_avg_usage} GB")
  print("Number of powered-on VMs:", count_powered_on_vms)
  print("Number of powered-on Windows VMs:", count_powered_on_windows_vms)
  print("Total Provisioned Disk Size for Powered-On VMs- TB:", disk_intb)
  print()

# List of file paths
file_paths = [
    '/Users/kay/Downloads/rv_sf.xlsx',]

for file_path in file_paths:
  process_excel(file_path)
