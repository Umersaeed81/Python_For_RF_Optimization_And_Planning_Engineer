# Extracting and Consolidating Specific Data from Multiple Excel Files Using Python


```python
import os
import pandas as pd
from openpyxl import load_workbook
```


```python
# Folder containing the Excel files
folder_path = "D:/Advance_Data_Sets/WCMS"  
```


```python
# Define the keys to extract
keys_to_extract = [
    "Site ID", "Region", "City", "District", "Location Name", "Site Address", "Latitude",
    "Longitude", "Zain ID", "Site Name 2G", "Site Name 3G", "Site Name LTE", "Site Name 5G",
    "eNodeB ID", "5GNodeB ID", "Project ", "Objective", "Band", "Site Type", "Building Height", "Antenna Type"
]
```


```python
# Initialize a list to store the data from all files
all_data = []

# Loop through all Excel files in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith(".xlsx"):  # Process only Excel files
        file_path = os.path.join(folder_path, file_name)
        
        # Load the workbook
        wb = load_workbook(file_path, data_only=True)
        sheet = wb.active
        
        # Initialize a dictionary to store the extracted data
        data = {"File Name": file_name}  # Include file name for reference
        
        # Loop through rows to find the desired keys and their values
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value in keys_to_extract:  # Check if the cell value is one of the keys
                    data[cell.value] = cell.offset(column=2).value  # Get the value from the adjacent cell
        
        # Append the dictionary to the list
        all_data.append(data)

# Convert the list of dictionaries to a DataFrame
df = pd.DataFrame(all_data)
```


```python
# Requierd Order
df1= df[['Site ID','Region','City','District','Location Name','Site Address',
        'Latitude','Longitude','Zain ID','Site Name 2G','Site Name 3G',
        'Site Name LTE','Site Name 5G','eNodeB ID','5GNodeB ID',
        'Project ','Objective','Band','Site Type',
        'Building Height','Antenna Type']]
```


```python
# Remove leading/trailing spaces and normalize spaces in between
df1.columns = df1.columns.str.strip().str.replace(r'\s+', ' ', regex=True)
```


```python
# Export Ouput
df1.to_excel("output.xlsx", index=False)
```
