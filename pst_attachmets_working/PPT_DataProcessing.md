## Folder Planning


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
    'D:/Advance_Data_Sets/UMTS_SunSet/PPT/City_Lelvel_Traffic']

# Process each folder
for folder_path in folder_paths:
    create_or_clear_folder(folder_path)
%reset -f
```

## Import Requried Libraries


```python
import os                         # 📁 Handle file system operations  
import zipfile                    # 📦 Work with ZIP files  
import numpy as np                # 🔢 Numerical computations  
import pandas as pd               # 📊 Data manipulation  
from pathlib import Path          # 🛤️ Handle file paths  
```

## Set input File Path


```python
 # 📌 Set the directory containing ZIP files  
zip_folder = "D:/Advance_Data_Sets/UMTS_SunSet/PPT" 
# 📄 Get a list of all ZIP files in the specified folder  
zip_files = list(Path(zip_folder).glob("*.zip")) 
```

## User Define Funcation


```python
def extract_csv_from_zip(zip_files, identifier, required_columns, rename_map=None):
    """
    Extracts and optionally renames columns in CSVs inside ZIP files.

    Parameters:
    - zip_files (list): List of ZIP file paths.
    - identifier (str): Substring to identify CSV files.
    - required_columns (list): List of original column names to extract.
    - rename_map (dict, optional): Mapping to rename columns.

    Returns:
    - pd.DataFrame: Combined DataFrame of extracted data.
    """
    dfs = []

    for zip_file in zip_files:
        with zipfile.ZipFile(zip_file, 'r') as z:
            csv_files = [f for f in z.namelist() if f"({identifier})" in f and f.endswith(".csv")]

            for csv_file in csv_files:
                with z.open(csv_file) as f:
                    try:
                        df = pd.read_csv(f, usecols=required_columns, parse_dates=["Date"],
                                         skiprows=range(6), skipfooter=1, engine='python',
                                         na_values=['NIL', '/0'])
                        if rename_map:
                            df = df.rename(columns=rename_map)

                        dfs.append(df)
                    except ValueError:
                        print(f"Skipping {csv_file} in {zip_file} - Missing required columns.")

    if dfs:
        return pd.concat(dfs, ignore_index=True)
    else:
        print(f"No matching CSV files found for identifier: {identifier}")
        return pd.DataFrame()
```

## Import Input Files

### 1) 🎯 Define identifiers and columns


```python
# 🎯 Define identifiers and columns
csv_info = {
    "2G_CG_DA": ['Date', 'GCell Group', 'GlobelTraffic'],
    "3G_CG_DA": ['Date','UCell Group', 'VS.AMR.Erlang.BestCell_SUM',	'Data Volume (GB)'],
    "4G_CG_DA": ['Date', 'LTE Cell Group', 'Data Volume (GB)',	'L.Traffic.User.VoIP.Avg',	'DL User Thrp (Mbps)']
}
```

### 2) Call User define funcation to import input Data


```python
# 📊 Extract all types of datasets
df_2g = extract_csv_from_zip(zip_files, "2G_CG_DA", csv_info["2G_CG_DA"]).\
            rename(columns={'GlobelTraffic': '2G_CS_Traffic'})
df_3g = extract_csv_from_zip(zip_files, "3G_CG_DA", csv_info["3G_CG_DA"]).\
            rename(columns={'VS.AMR.Erlang.BestCell_SUM': '3G_CS_Traffic','Data Volume (GB)': '3G_DataVolume_GB'})
df_4g = extract_csv_from_zip(zip_files, "4G_CG_DA", csv_info["4G_CG_DA"]).\
            rename(columns={'L.Traffic.User.VoIP.Avg': '4G_CS_Traffic','Data Volume (GB)': '4G_DataVolume_GB'}).fillna(0)
```

## Data Processing

### 1) City Name Extraction


```python
df_2g['City'] = df_2g['GCell Group'].str.replace('2G_', '', regex=False)
df_3g['City'] = df_3g['UCell Group'].str.replace('3G_', '', regex=False)
df_4g['City'] = df_4g['LTE Cell Group'].str.replace('4G_', '', regex=False)
```

### 2) Per Data CS and PS Traffic - City Level


```python
df_merge = pd.merge(df_2g,df_3g,on=['Date','City'],how='left').merge(df_4g,on=['Date','City'],how='left')
```

### 3) Re-Order the column List


```python
df_merge1 =  df_merge[['Date','City','2G_CS_Traffic','3G_CS_Traffic','3G_DataVolume_GB','4G_DataVolume_GB','4G_CS_Traffic']]
```

### 4) Calcualte Total CS - PS Traffic City & Day Level


```python
df_merge1 = df_merge1.copy(deep=True)
df_merge1['Total_CS_Traffic'] = (df_merge1['2G_CS_Traffic'] + df_merge1['3G_CS_Traffic'] + df_merge1['4G_CS_Traffic'])
df_merge1['Total_Data_volume'] = (df_merge1['3G_DataVolume_GB'] + df_merge1['4G_DataVolume_GB'])
```

## Export CS & PS Traffic City & Day Level


```python
path = 'D:/Advance_Data_Sets/UMTS_SunSet/PPT/City_Lelvel_Traffic'
os.chdir(path)

# # Define groups of metrics
# cs_traffic_cols = ['2G_CS_Traffic', '3G_CS_Traffic', '4G_CS_Traffic','Total_CS_Traffic']
# data_volume_cols = ['3G_DataVolume_GB','4G_DataVolume_GB','Total_Data_volume']

# df_merge1 = df_merge1.copy()
# df_merge1['Date'] = pd.to_datetime(df_merge1['Date'])

# # Export each city to its own Excel file
# for city in df_merge1['City'].unique():
#     city_df = df_merge1[df_merge1['City'] == city]
#     file_name = f"{city}_Metrics_Report.xlsx"
    
#     with pd.ExcelWriter(file_name) as writer:
#         # Sheet for CS Traffic
#         cs_df = city_df[['Date'] + cs_traffic_cols]
#         cs_df.to_excel(writer, sheet_name='CS_Traffic', index=False,datetime_format='dd/mm/yyyy')
        
#         # Sheet for Data Volume
#         data_df = city_df[['Date'] + data_volume_cols]
#         data_df.to_excel(writer, sheet_name='Data_Volume', index=False,datetime_format='dd/mm/yyyy')




# Define groups of metrics
cs_traffic_cols = ['2G_CS_Traffic', '3G_CS_Traffic', '4G_CS_Traffic','Total_CS_Traffic']
data_volume_cols = ['3G_DataVolume_GB','4G_DataVolume_GB','Total_Data_volume']

df_merge1 = df_merge1.copy()
df_merge1['Date'] = pd.to_datetime(df_merge1['Date'])

# Export each city to its own Excel file
for city in df_merge1['City'].unique():
    city_df = df_merge1[df_merge1['City'] == city]
    file_name = f"{city}_Metrics_Report.xlsx"
    
    with pd.ExcelWriter(file_name, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
        # Sheet for CS Traffic
        cs_df = city_df[['Date'] + cs_traffic_cols]
        cs_df.to_excel(writer, sheet_name='CS_Traffic', index=False)
        
        # Sheet for Data Volume
        data_df = city_df[['Date'] + data_volume_cols]
        data_df.to_excel(writer, sheet_name='Data_Volume', index=False)

```

## LTE Data Pre-Processing

### 1) Convet GB in TB


```python
df_4g['4G_DataVolume_TB'] = (df_4g['4G_DataVolume_GB']/1024).round(1)
```

### 2) Identify City and Its Technology


```python
# Extract text before '_Target' or '_Control'
df_4g['First_Column'] = df_4g['LTE Cell Group'].str.extract(r'^(.*?)_(?:Target|Control)')
# Extract after the last underscore
df_4g['Second_Column'] = df_4g['LTE Cell Group'].str.extract(r'_([^_]+)$')
```

### 3) LTE City Level (Data Volume) - Reshpare Layer Level


```python
df_4g_dvtb = df_4g.pivot_table(index=['Date','First_Column'],columns='Second_Column',values='4G_DataVolume_TB').reset_index()
```

### 4) Re-Order Colums


```python
df_4g_dvtb1 = df_4g_dvtb[['First_Column','Date','L1800','L2100','L900','City']]
```

### 5) LTE City Level (DL User Thrp) - Reshpare Layer Level


```python
df_4g_dvtp = df_4g.pivot_table(index=['Date','First_Column'],columns='Second_Column',values='DL User Thrp (Mbps)').round(1).reset_index()
```

### 6) Re-Order Columns


```python
df_4g_dvtp1 = df_4g_dvtp[['First_Column','Date','City','L1800','L2100','L900']]
```

## Export 4G_Data_Volume_Report


```python
path = 'D:/Advance_Data_Sets/UMTS_SunSet/PPT/City_Lelvel_Traffic'
os.chdir(path)

df_4g_dvtp1 = df_4g_dvtp1.copy()
df_4g_dvtp1['Date'] = pd.to_datetime(df_4g_dvtp1['Date'])

# Create an Excel writer
with pd.ExcelWriter("4G_thoroughput_Report.xlsx", engine='xlsxwriter',datetime_format='dd/mm/yyyy') as writer:
    for name in df_4g_dvtp1['First_Column'].unique():
        subset = df_4g_dvtp1[df_4g_dvtp1['First_Column'] == name]
        subset.to_excel(writer, sheet_name=str(name)[:31], index=False)
```

## Export 4G_Trhoughput


```python
path = 'D:/Advance_Data_Sets/UMTS_SunSet/PPT/City_Lelvel_Traffic'
os.chdir(path)

df_4g_dvtb1 = df_4g_dvtb1.copy()
df_4g_dvtb1 ['Date'] = pd.to_datetime(df_4g_dvtb1['Date'])


# Create an Excel writer
with pd.ExcelWriter("4G_Data_Volume_Report.xlsx", engine='xlsxwriter',datetime_format='dd/mm/yyyy') as writer:
    for name in df_4g_dvtb1['First_Column'].unique():
        subset = df_4g_dvtb1[df_4g_dvtb1['First_Column'] == name]
        subset.to_excel(writer, sheet_name=str(name)[:31], index=False)
```


```python
%reset -f
```


```python

```
