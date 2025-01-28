#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

-----------------------------------------------------
# Automating Data Extraction from Excel Files in Python: A Step-by-Step Guide

## Introduction
Managing data from multiple Excel files can be a tedious task, especially when dealing with large datasets that require extracting specific details. This article walks you through a Python script designed to automate the process of extracting key information from Excel files stored in a folder. By leveraging Python libraries like `os`, `pandas`, and `openpyxl`, this solution efficiently extracts and organizes the data into a consolidated output.

## Objectives of the Script
The goal of this script is to:

1. Extract specific key-value pairs from multiple Excel files stored in a folder.
2. Normalize the data into a consistent format.
3. Save the extracted information into a well-organized output file.

## Step-by-Step Explanation of the Code

## 1. Setting Up the Environment
The script starts by importing the necessary libraries:

- `os` for file operations.
- `pandas` for data manipulation and exporting.
- `openpyxl` to interact with Excel files.

The folder containing the Excel files is specified with `folder_path`.

## Import Required Libraries
```python
import os
import pandas as pd
from openpyxl import load_workbook
```

## Set Excel Files Path 

```python
# Folder containing the Excel files
folder_path = "D:/Advance_Data_Sets/WCMS"  
```

## Define the keys to extract
```python
# Define the keys to extract
keys_to_extract = [
    "Site ID", "Region", "City", "District", "Location Name", "Site Address", "Latitude",
    "Longitude", "Zain ID", "Site Name 2G", "Site Name 3G", "Site Name LTE", "Site Name 5G",
    "eNodeB ID", "5GNodeB ID", "Project ", "Objective", "Band", "Site Type", "Building Height", "Antenna Type"
]
```

## Iterating Through Excel Files to Extract Key-Value Data with Python

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

## Re-order the requied columns
```python
# Requierd Order
df1= df[['Site ID','Region','City','District','Location Name','Site Address',
        'Latitude','Longitude','Zain ID','Site Name 2G','Site Name 3G',
        'Site Name LTE','Site Name 5G','eNodeB ID','5GNodeB ID',
        'Project ','Objective','Band','Site Type',
        'Building Height','Antenna Type']]
```

## Pre-Processing
```python
# Remove leading/trailing spaces and normalize spaces in between
df1.columns = df1.columns.str.strip().str.replace(r'\s+', ' ', regex=True)
```

## Export Output
```python
# Export Ouput
df1.to_excel("output.xlsx", index=False)
```
