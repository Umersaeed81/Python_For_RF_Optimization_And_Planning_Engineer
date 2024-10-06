#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

## Import Required Libraries


```python
import os
import shutil
import pandas as pd
```

## Delete Files From the Folder, If Folder Not Found Created the Folder


```python
# Folder path
folder_path = r"D:/Advance_Data_Sets/GUL/GUL_Output"

# Check if the folder exists
if os.path.exists(folder_path):
    # Iterate over the files in the folder and delete them
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted file: {file_path}")
        except Exception as e:
            print(f"Error deleting {file_path}: {e}")
else:
    # If the folder doesn't exist, create it
    print(f"Folder {folder_path} does not exist. Creating...")
    os.makedirs(folder_path)
    print(f"Folder {folder_path} created.")
```

## Move Daily KPIs and RF Export in Required Folders


```python
#  Folder/File Path in tupl format
df1 = [('GUL_Utilization_KPIs_NW_UMTS',
        'D:/Advance_Data_Sets/GUL/3G',
        'E:/test_path'),
       
        ('GUL_Utilization_KPIs_NW_LTE_',
        'D:/Advance_Data_Sets/GUL/4G',
        'E:/test_path')]

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

## Re-set Variable


```python
# re-set all the variable from the RAM
%reset -f
```
