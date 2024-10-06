#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# 0. Delete RF Export Files From the System


```python
import os
import shutil

def delete_all_files_and_folders(folder_path):
    # Ensure the folder path exists before proceeding
    if os.path.exists(folder_path):
        # Iterate over the contents of the folder
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)

            # If it's a file, remove it
            if os.path.isfile(item_path):
                os.remove(item_path)

            # If it's a directory, remove it and its contents recursively
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
    else:
        print(f"The folder path '{folder_path}' does not exist.")
        return False
    return True

# List of folder paths
folder_paths = ["D:/Advance_Data_Sets/License/LTE_NodeB_Level",
                "D:/Advance_Data_Sets/RF_Export/GSM/00_Frequency_Export",
                "D:/Advance_Data_Sets/RF_Export/GSM/01_NBR_Export",
                "D:/Advance_Data_Sets/RF_Export/GSM/03_GTRX_Export",
                "D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export",
                "D:/Advance_Data_Sets/RF_Export/LTE/Cell",
                "D:/Advance_Data_Sets/RF_Export/UMTS/00_Cell_Export",
                "D:/Advance_Data_Sets/RF_Export/UMTS/01_NBR_Export",
                "D:/Advance_Data_Sets/RF_Export/GSM/01_NBR_Export",
                "D:/Advance_Data_Sets/RF_Export/GSM/03_GTRX_Export",
                "D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export",
                "D:/Advance_Data_Sets/Output_Folder",
                "D:/Advance_Data_Sets/GSM_Utilization_Output_Files",
                "D:/Advance_Data_Sets/Pagging/Daily",
                "D:/Advance_Data_Sets/M2_Report",
                "D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files",
                "D:/Advance_Data_Sets/TXN_DataSets/GSM_CGID_log_Analysis"]  
# Initialize counters
exist_count = 0
not_exist_count = 0

# Delete all files and folders in each path and count them
for path in folder_paths:
    if delete_all_files_and_folders(path):
        exist_count += 1
    else:
        not_exist_count += 1

# Print the counts
print("Summary:")
print(f" - Number of existing paths: {exist_count}")
print(f" - Number of non-existing paths: {not_exist_count}")
# re-set all the variable from the RAM
%reset -f
```

## 1. Folder Planning


```python
import os

def create_or_clear_folder(folder_path):
    # Check if the folder exists
    if os.path.exists(folder_path):
        print(f"Removing contents of existing folder '{folder_path}'...")
        
        # Remove all files in the folder
        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    os.rmdir(file_path)
            except Exception as e:
                print(f"Error: {e}")

        print("Contents removed successfully.")
    else:
        print(f"Creating folder '{folder_path}'...")
        os.makedirs(folder_path)
        print("Folder created successfully.")

    print("Task completed.")

# Specify folder paths
folder_paths = [
    'D:\Advance_Data_Sets\Output_Folder\M2_GSM',
    'D:\Advance_Data_Sets\Output_Folder\M2_LTE',
    'D:\Advance_Data_Sets\Output_Folder\M2_UMTS'
]

# Process each folder
for folder_path in folder_paths:
    create_or_clear_folder(folder_path)
%reset -f
```

# 2. Move Hourly KPIs in Sub-Folders


```python
# import the libraries
import os
import shutil
import pandas as pd

# Define the base dates
base_dates = ['22092024']

# Initialize an empty list to store the tuples
df1 = []

# Define the file names and paths
file_info = [
    ('UMTS_CE_Utilization', 
     'Congestion/CE_Utilization/', 
     'Congestion/CE_Utilization/12_Dec_2023'),
    
    ('GSM_Cell_Houlry', 
     'KPIs_Analysis/Hourly_KPIs_Cell_Level/GSM/', 
     'KPIs_Analysis/Hourly_KPIs_Cell_Level/GSM/Dec'),
    
    ('LTE_Cell_Hourly', 
     'KPIs_Analysis/Hourly_KPIs_Cell_Level/LTE/',
     'KPIs_Analysis/Hourly_KPIs_Cell_Level/LTE/12_Dec_2023'),
    
    ('UMTS_Cell_Hourly', 
     'KPIs_Analysis/Hourly_KPIs_Cell_Level/UMTS', 
     'KPIs_Analysis/Hourly_KPIs_Cell_Level/UMTS/12_Dec_2023'),
    
    ('GSM_BTS_Hourly', 
     'TXN_DataSets/GSM_BTS_Hourly/', 
     'TXN_DataSets/GSM_BTS_Hourly/12_Dec'),
    
    ('LTE_Hourly_BTS_Level', 
     'TXN_DataSets/LTE_TNL_BTS_Hourly/', 
     'TXN_DataSets/LTE_TNL_BTS_Hourly/12_Dec_2023'),
    
    ('LTE_Nationwide_RX', 
     'TXN_DataSets/LTE_Transmission_Link_Capacity/', 
     'TXN_DataSets/LTE_Transmission_Link_Capacity/12_Dec'),
    
    ('UMTS_BTS_Hourly', 
     'TXN_DataSets/UMTS_Hourly_BTS/', 
     'TXN_DataSets/UMTS_Hourly_BTS/12_Dec'),
    
    ('IPPath_Hourly', 
     'TXN_DataSets/UMTS_Transmission_Hourly/IPPATH_Hourly/', 
     'TXN_DataSets/UMTS_Transmission_Hourly/IPPATH_Hourly/12_Dec_2023'),
    
    ('IPPool_Hourly', 
     'TXN_DataSets/UMTS_Transmission_Hourly/IPPOOL_Hourly/', 
     'TXN_DataSets/UMTS_Transmission_Hourly/IPPOOL_Hourly/12_Dec_2023')
]

# Loop through the dates and create tuples
for date in base_dates:
    for file_name, sr_path, target_path in file_info:
        file_name = '{}_{}.zip'.format(file_name, date)
        sr_path = 'D:/Advance_Data_Sets/{}'.format(sr_path)
        target_path = 'D:/Advance_Data_Sets/{}'.format(target_path)
        df1.append((file_name, sr_path, target_path))

# Create a DataFrame
df2 = pd.DataFrame(df1, columns=['File', 'Sr_Path', 'Target_Path'])


# Initialize counters
success_count = 0
failure_count = 0

# Iterate through the rows
for index, row in df2.iterrows():
    source_file = os.path.join(row['Sr_Path'], row['File'])
    target_path = row['Target_Path']
    
    try:
        # Check if the source file exists
        if os.path.exists(source_file):
            # Create the target directory if it doesn't exist
            os.makedirs(target_path, exist_ok=True)
            
            # Move the file
            shutil.move(source_file, target_path)
            print(f"The file '{source_file}' was successfully moved.")
            success_count += 1
        else:
            print(f"The file '{source_file}' does not exist.")
            failure_count += 1
    except Exception as e:
        print(f"An error occurred while moving the file '{source_file}': {e}")
        failure_count += 1

# Print summary
print(f"Summary:")
print(f"  - {success_count} files successfully moved")
print(f"  - {failure_count} files failed to move.")

#re-set all the variable from the RAM
%reset -f
```

# 3. Delete DA and BH KPIs


```python
# import requied libraries
import os
import shutil
import pandas as pd

# set the date before run the code
date = '_02102024.zip'

#  Folder/File Path in tupl format
df1 = [
#         ('D:/Advance_Data_Sets/HW_Scripts/LTE_Main_Diversity/LTE_Main_Diversity{}'.format(date)),
#         ('D:/Advance_Data_Sets/HW_Scripts/UMTS_Main_Diversity/UMTS_Main_Diversity{}'.format(date)),
        ('D:/Advance_Data_Sets/KPIs_Analysis/BH_KPIs_Cell_Level/GSM/GSM_Cell_BH{}'.format(date)),
        ('D:/Advance_Data_Sets/KPIs_Analysis/BH_KPIs_Cell_Level/UMTS/UMTS_Cell_BH{}'.format(date)),
#          ('D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/GSM/GSM_Cell_DA{}'.format(date)),
#          ('D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/LTE/LTE_Cell_DA{}'.format(date)),
#          ('D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/UMTS/UMTS_Cell_DA{}'.format(date)),
        ('D:/Advance_Data_Sets/KPIs_Analysis/Inter_BSC_HSR/KPIs/INTER_BSC_HSR{}'.format(date)),
        ('D:/Advance_Data_Sets/SLA/Cluster_DA_KPIs/GSM_Cluster_DA{}'.format(date)),
        ('D:/Advance_Data_Sets/SLA/Cluster_BH_KPIs/GSM_Cluster_BH{}'.format(date)),
        ('D:/Advance_Data_Sets/TA_Analysis/GSM/DA/TA_Analysis_DA{}'.format(date)),
        ('D:/Advance_Data_Sets/TA_Analysis/GSM/Hourly/TA_Analysis_Hourly{}'.format(date)),
        ('D:/Advance_Data_Sets/TXN_DataSets/GSM_BTS_DA/GSM_BTS_DA{}'.format(date)),
        ('D:/Advance_Data_Sets/TXN_DataSets/LTE_TNL_BTS_DA/LTE_BTS_DA{}'.format(date)),
        ('D:/Advance_Data_Sets/TXN_DataSets/UMTS_DA_BTS/UMTS_BTS_DA{}'.format(date)),
#        ('D:/Advance_Data_Sets/TXN_DataSets/UMTS_Transmission_DA/IPPATH_DA/IPPath_DA{}'.format(date)),
#        ('D:/Advance_Data_Sets/TXN_DataSets/UMTS_Transmission_DA/IPPOOL_DA/IPPool_DA{}'.format(date)),       
#        ('D:/Advance_Data_Sets/KPIs_Analysis/BH_KPIs_Cell_Level/UMTS_RBH/UMTS_Cell_RBH{}'.format(date)),
#        ('D:/Advance_Data_Sets/KPIs_Analysis/BH_KPIs_Cell_Level/LTE_RBH/LTE_Cell_RBH{}'.format(date))  
      ]

# Create a DataFrame
df2 = pd.DataFrame(df1, columns=['File'])

# Initialize counters
deleted_successfully = 0
not_found = 0

# Iterate through the rows and delete the files
for index, row in df2.iterrows():
    file_path = row['File']
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f"{file_path} deleted successfully.")
        deleted_successfully += 1
    else:
        print(f"{file_path} does not exist.")
        not_found += 1

# Print the counts
print('Summary is given below;')
print(f" - Deleted successfully: {deleted_successfully} files")
print(f" - Not found: {not_found} files")

#re-set all the variable from the RAM
%reset -f
```

# 4. Move Daily KPIs and RF Export in Required Folders


```python
# import required libraries
import os
import shutil
import pandas as pd

#  Folder/File Path in tupl format
df1 = [('UMTS_CE_Utilization_',
        'D:/Advance_Data_Sets/Congestion/CE_Utilization',
        'E:/test_path'),
       
        ('UMTS_Main_Diversity_',
        'D:/Advance_Data_Sets/HW_Scripts/UMTS_Main_Diversity',
        'E:/test_path'),
       
        ('LTE_Main_Diversity_',
        'D:/Advance_Data_Sets/HW_Scripts/LTE_Main_Diversity',
        'E:/test_path'),
      
        ('GSM_Cell_BH_',
        'D:/Advance_Data_Sets/KPIs_Analysis/BH_KPIs_Cell_Level/GSM',
        'E:/test_path'),
       
        ('UMTS_Cell_BH_',
        'D:/Advance_Data_Sets/KPIs_Analysis/BH_KPIs_Cell_Level/UMTS',
        'E:/test_path'),
       
        ('GSM_Cell_DA_',
        'D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/GSM',
        'E:/test_path'),
       
        ('UMTS_Cell_DA_',
        'D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/UMTS',
        'E:/test_path'),
       
        ('LTE_Cell_DA_',
        'D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/LTE',
        'E:/test_path'),
           
        ('GSM_Cell_Houlry_',
        'D:/Advance_Data_Sets/KPIs_Analysis/Hourly_KPIs_Cell_Level/GSM',
        'E:/test_path'),
       
        ('UMTS_Cell_Hourly_',
        'D:/Advance_Data_Sets/KPIs_Analysis/Hourly_KPIs_Cell_Level/UMTS',
        'E:/test_path'),
       
       ('LTE_Cell_Hourly_',
        'D:/Advance_Data_Sets/KPIs_Analysis/Hourly_KPIs_Cell_Level/LTE',
        'E:/test_path'),
       
        ('INTER_BSC_HSR_',
        'D:/Advance_Data_Sets/KPIs_Analysis/Inter_BSC_HSR/KPIs',
        'E:/test_path'),
       
        ('2GCellParamExport',
        'D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export',
        'E:/test_path'),
       
        ('2GCellFreqExport',
        'D:/Advance_Data_Sets/RF_Export/GSM/00_Frequency_Export',
        'E:/test_path'),
       
        ('2GRnpExport',
        'D:/Advance_Data_Sets/RF_Export/GSM/01_NBR_Export',
        'E:/test_path'),
       
        ('2GTrxExport',
        'D:/Advance_Data_Sets/RF_Export/GSM/03_GTRX_Export',
        'E:/test_path'),
       
        ('3GCellParamExport',
        'D:/Advance_Data_Sets/RF_Export/UMTS/00_Cell_Export',
        'E:/test_path'),
       
        ('3GRnpExport',
        'D:/Advance_Data_Sets/RF_Export/UMTS/01_NBR_Export',
        'E:/test_path'),
       
        ('wetransfer_configurationdata',
        'D:/Advance_Data_Sets/RF_Export/LTE/Cell',
        'E:/test_path'),
       
        ('ConfigurationData_',
        'D:/Advance_Data_Sets/RF_Export/LTE/Cell',
        'E:/test_path'),
       
       
        ('LTE License',
        'D:/Advance_Data_Sets/License/LTE_NodeB_Level',
        'E:/test_path'),
       
       
        ('GSM_Cluster_BH_',
        'D:/Advance_Data_Sets/SLA/Cluster_BH_KPIs',
        'E:/test_path'),
       
         ('GSM_Cluster_DA_',
        'D:/Advance_Data_Sets/SLA/Cluster_DA_KPIs',
        'E:/test_path'),
       
         ('TA_Analysis_DA_',
        'D:/Advance_Data_Sets/TA_Analysis/GSM/DA',
        'E:/test_path'),
       
         ('TA_Analysis_Hourly_',
        'D:/Advance_Data_Sets/TA_Analysis/GSM/Hourly',
        'E:/test_path'),
       
         ('GSM_BTS_DA_',
        'D:/Advance_Data_Sets/TXN_DataSets/GSM_BTS_DA',
        'E:/test_path'),
    
        ('UMTS_BTS_DA_',
        'D:/Advance_Data_Sets/TXN_DataSets/UMTS_DA_BTS',
        'E:/test_path'),
       
       
        ('LTE_BTS_DA_',
        'D:/Advance_Data_Sets/TXN_DataSets/LTE_TNL_BTS_DA',
        'E:/test_path'),
       
       
        ('GSM_BTS_Hourly_',
        'D:/Advance_Data_Sets/TXN_DataSets/GSM_BTS_Hourly',
        'E:/test_path'),
       
        ('UMTS_BTS_Hourly_',
        'D:/Advance_Data_Sets/TXN_DataSets/UMTS_Hourly_BTS',
        'E:/test_path'),
              
       ('LTE_Hourly_BTS_Level_',
        'D:/Advance_Data_Sets/TXN_DataSets/LTE_TNL_BTS_Hourly',
        'E:/test_path'),
      
        ('LTE_Nationwide_RX_',
        'D:/Advance_Data_Sets/TXN_DataSets/LTE_Transmission_Link_Capacity',
        'E:/test_path'),
       
        ('IPPath_DA_',
        'D:/Advance_Data_Sets/TXN_DataSets/UMTS_Transmission_DA/IPPATH_DA',
        'E:/test_path'),
       
        ('IPPool_DA_',
        'D:/Advance_Data_Sets/TXN_DataSets/UMTS_Transmission_DA/IPPOOL_DA',
        'E:/test_path'),

        ('IPPath_Hourly_',
        'D:/Advance_Data_Sets/TXN_DataSets/UMTS_Transmission_Hourly/IPPATH_Hourly',
        'E:/test_path'),
       
        ('IPPool_Hourly_',
        'D:/Advance_Data_Sets/TXN_DataSets/UMTS_Transmission_Hourly/IPPOOL_Hourly',
        'E:/test_path'),
       
        ('CDIG',
        'D:/Advance_Data_Sets/TXN_DataSets/GSM_CGID_log_Analysis',
        'E:/test_path'),
       
        ('UMTS_Cell_RBH',
        'D:/Advance_Data_Sets/KPIs_Analysis/BH_KPIs_Cell_Level/UMTS_RBH',
        'E:/test_path'),
       
        ('LTE_Cell_RBH',
        'D:/Advance_Data_Sets/KPIs_Analysis/BH_KPIs_Cell_Level/LTE_RBH',
        'E:/test_path'),
       
        ('NSS',
        'D:/Advance_Data_Sets/Pagging/Daily',
        'E:/test_path'),
      ]

# Create a DataFrame
df2 = pd.DataFrame(df1, columns=['File','Target_Folder','Sr_Folder'])


# Initialize counters
prefix_found_and_moved = 0
prefix_not_found = 0

# Iterate through the rows in the DataFrame
for index, row in df2.iterrows():
    source_folder = row['Sr_Folder']
    target_folder = row['Target_Folder']

    # Check if source folder exists
    if not os.path.exists(source_folder):
        print(f"Source folder '{source_folder}' does not exist.")
        continue

    # Check if target folder exists, create it if it doesn't
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)

    # Check if any file with the specified prefix exists in the source folder
    files_with_prefix = [file for file in os.listdir(source_folder) if file.startswith(row['File'])]
    if not files_with_prefix:
        print(f"No files with prefix '{row['File']}' found in '{source_folder}'.")
        prefix_not_found += 1
        continue

    # Count the number of files with the specified prefix
    num_files_with_prefix = len(files_with_prefix)
    print(f"Number of files with prefix '{row['File']}' found in '{source_folder}': {num_files_with_prefix}")

    for file_name in files_with_prefix:
        source_path = os.path.join(source_folder, file_name)
        target_path = os.path.join(target_folder, file_name)
        shutil.move(source_path, target_path)

        # Print the file names being moved
        print(f"Moving '{file_name}' from '{source_folder}' to '{target_folder}'")

    prefix_found_and_moved += 1

# Print summary
print(f"Summary:")
print(f"  - Prefixes found and moved: {prefix_found_and_moved}")
print(f"  - Prefixes not found: {prefix_not_found}")
```

# 5. Unzip LTE RF Export


```python
import os
import zipfile
# Set RF Export File Path
working_directory = 'D:/Advance_Data_Sets/RF_Export/LTE/Cell'
os.chdir(working_directory)

# unzip files here
for file in os.listdir(working_directory):
    if zipfile.is_zipfile(os.path.join(working_directory, file)):
        with zipfile.ZipFile(os.path.join(working_directory, file), 'r') as zip_ref:
            zip_ref.extractall(working_directory)

# delete zip file from the folder            
for filename in os.listdir(working_directory):
    if filename.endswith('.zip'):
        os.unlink(os.path.join(working_directory, filename))

# re-set all the variable from the RAM
%reset -f
```

## 5.1 Delete file base on size of the file


```python
import os
import zipfile
# Set RF Export File Path
working_directory = 'D:/Advance_Data_Sets/RF_Export/LTE/Cell'
os.chdir(working_directory)

# Minimum file size in bytes (50 KB)
min_file_size = 50 * 1024

# Iterate through all files in the folder
for filename in os.listdir(working_directory):
    if filename.endswith('.xlsx'):  # Assuming all files are in Excel format
        file_path = os.path.join(working_directory, filename)
        
        # Check the size of the file
        file_size = os.path.getsize(file_path)
        
        if file_size < min_file_size:
            # Delete the file
            os.remove(file_path)
            
            print(f"File '{filename}' deleted.")
        else:
            print(f"File '{filename}' not deleted. Size is {file_size} bytes.")

print("Process completed.")

# re-set all the variable from the RAM
%reset -f
```

# 6. Unzip(.tar.gz) CDIG File (Row File)


```python
import os
import gzip
import shutil
import tarfile

# Define the directory containing the .tar.gz file
directory_path = 'D:/Advance_Data_Sets/TXN_DataSets/GSM_CGID_log_Analysis/'
tar_file_name = os.listdir(directory_path)[0]  # Access the first file in the list

tar_file_path = os.path.join(directory_path, tar_file_name)

# Define the extraction directory
extraction_path = 'D:/Advance_Data_Sets/TXN_DataSets/GSM_CGID_log_Analysis/'

# Create the extraction directory if it doesn't exist
os.makedirs(extraction_path, exist_ok=True)

# Extract the .tar.gz file
with open(tar_file_path, 'rb') as f_in, \
        gzip.open(f_in, 'rb') as f_out_tar, \
        tarfile.open(fileobj=f_out_tar, mode='r|*') as tar:
    tar.extractall(path=extraction_path)
```

## 6.1 Delete Unwanted Files


```python
# delete tar.gz file from the folder            
for filename in os.listdir(directory_path):
    if filename.endswith('.tar.gz'):
        os.unlink(os.path.join(directory_path, filename))
```

## 6.2 Move CDIG Log from Sub Folder to Main Folder


```python
import os
import shutil

# Define the main directory path
main_directory = 'D:/Advance_Data_Sets/TXN_DataSets/GSM_CGID_log_Analysis'

# Iterate over all subdirectories recursively
for root, dirs, files in os.walk(main_directory):
    for subfolder_name in dirs[:]:  # Make a copy of 'dirs' to iterate over
        subfolder_path = os.path.join(root, subfolder_name)
        # Check if it's a subdirectory
        if os.path.isdir(subfolder_path) and subfolder_name != os.path.basename(main_directory):
            # Move the files from the subdirectory to the main directory
            for item in os.listdir(subfolder_path):
                item_path = os.path.join(subfolder_path, item)
                shutil.move(item_path, main_directory)
            # Remove the empty subdirectory
            os.rmdir(subfolder_path)
```
