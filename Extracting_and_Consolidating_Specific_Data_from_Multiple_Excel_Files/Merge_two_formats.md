#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

-----------------------------------------------------

## 1. Setting Up the Environment


```python
import os
import pandas as pd
from openpyxl import load_workbook
```

## 2. Set Excel Files Path 


```python
# Folder containing the Excel files
folder_path = "D:/Advance_Data_Sets/WCMS" 
os.chdir(folder_path)
```

## 3. Defining Keys to Extract


```python
keys_to_extract = [
    "Site ID", "Region", "City", "District", "Location Name", "Site Address", "Latitude",
    "Longitude", "Zain ID", "Site Name 2G", "Site Name 3G", "Site Name LTE", "Site Name 5G",
    "eNodeB ID", "5GNodeB ID", "Project ", "Objective", "Band", "Site Type", "Building Height",
    "Antenna Type"
]
```

## 4. Processing Files and Extracting Data


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
        
        # First, check if "Site ID" has the value "Rev:" before deciding which logic to use
        use_code_2 = False
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value == "Site ID":
                    site_id_value = cell.offset(column=2).value  # Check Site ID value from Code-1 logic
                    if isinstance(site_id_value, str) and site_id_value.strip() == "Rev:":
                        use_code_2 = True
                    break
            if use_code_2:
                break
        
        # Extract data based on the selected logic
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value in keys_to_extract:  # Check if the cell value is one of the keys
                    if use_code_2:
                        # Use Code-2 logic
                        if cell.value == "Antenna Type":
                            data[cell.value] = cell.offset(column=2).value  # Use column offset 2
                        else:
                            data[cell.value] = cell.offset(column=1).value  # Use column offset 1 for others
                    else:
                        # Use Code-1 logic
                        data[cell.value] = cell.offset(column=2).value  # Get the value from the adjacent cell
        
        # Append the dictionary to the list
        all_data.append(data)

# Convert the list of dictionaries to a DataFrame
df = pd.DataFrame(all_data)
```

## 5. Organizing the Columns


```python
# Required Order
df1 = df[['File Name','Site ID','Region','City','District','Location Name','Site Address',
          'Latitude','Longitude','Zain ID','Site Name 2G','Site Name 3G',
          'Site Name LTE','Site Name 5G','eNodeB ID','5GNodeB ID',
          'Project ','Objective','Band','Site Type',
          'Building Height','Antenna Type']]
```


```python
# Remove leading/trailing spaces and normalize spaces in between
df1.columns = df1.columns.str.strip().str.replace(r'\s+', ' ', regex=True)
```

## 6. Exporting the Final Output


```python
# Export Ouput
df1.to_excel("Output.xlsx", index=False)
```
