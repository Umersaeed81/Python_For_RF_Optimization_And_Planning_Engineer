#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# LTE NodB Level Licence Audit

## Import Required Libraries


```python
import os
import shutil
import zipfile
import patoolib
#pip install patool
import numpy as np
import pandas as pd
from glob import glob
from datetime import datetime, timedelta
import warnings
warnings.simplefilter("ignore")
```

## Set LTE NodB Level Licence Path


```python
# Set LTE NodB Level Licence Path
folder_path = 'D:/Advance_Data_Sets/License/LTE_NodeB_Level'
os.chdir(folder_path)
```

## Unzip rar Files


```python
# Get a list of all .rar files in the directory
rar_files = [file for file in os.listdir() if file.endswith(".rar")]

# Extract all .rar files
for rar_file in rar_files:
    patoolib.extract_archive(rar_file, outdir=".")
```

## List License Files in the Given Path


```python
df= sorted(glob('**/*.csv'))
```

## Import & Concat License Files


```python
df0=pd.concat((pd.read_csv(file,skiprows=range(9),
              skipfooter=1,engine='python',
              na_values=['NIL','/0']) for file in df)).\
              reset_index(drop=True)
```

## Delete Unzip (.rar in this case) Folder


```python
# Python Script for Clearing Subfolders in a Specific path
for subfolder in os.listdir(folder_path):
    subfolder_path = os.path.join(folder_path, subfolder)
    if os.path.isdir(subfolder_path):
        shutil.rmtree(subfolder_path)
```

## Export Combine Licence File


```python
df0.to_excel('LTE_NodeB_Level_Licence.xlsx',index=False,sheet_name="Licence")
```

## Filter Required License


```python
df1 = df0[df0['SBOM Description'].eq('RRC Connected User(FDD)')].\
        reset_index(drop=True).\
        rename(columns={'Node Name': 'eNodeB Name'})\
        [['eNodeB Name','SBOM Description','SBOM Model','Allocated']]
```

## Input (Max Users) File Path


```python
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/Max_Number_User'
os.chdir(folder_path)
```

## Filter Required Files (Max Users)


```python
# List all files in the path
file_list = glob('*.csv')

# Calculate the date for three days ago
required_date = datetime.now() - timedelta(days=7)
required_date_str = required_date.strftime("%d%m%Y")

# Filter files that match the required date or the two days before
filtered_file_list = [file for file in file_list if any((required_date + timedelta(days=i)).strftime("%d%m%Y") in file for i in range(7))]

if filtered_file_list:
    print("Filtered File List:", filtered_file_list)
else:
    print(f"No file found with the required dates in the last three days.")
```

## Import and concat Max Users Files


```python
df2=pd.concat((pd.read_csv(file,
        engine='python',na_values=['NIL','/0'],
        parse_dates=["Date"]) for file in filtered_file_list)).\
        sort_values('Date').reset_index(drop=True)
```

## Re-Shape (pivot_table) Max Users DataFrame


```python
df3= df2.pivot_table(index=['eNodeB Name'],
    columns="Date",
    values='Maximum User Number').reset_index()
```

## Merge Max Users Pivot_table with Filter License


```python
df4 = pd.merge(df3,df1,on=['eNodeB Name'],how='left')
```

## Convert Data Type of Allocated Column


```python
df4['Allocated'] = df4['Allocated'].fillna(0).astype(int)
```

## Conditional Count


```python
# Check if 'SBOM Model' column is not blank
mask = df4['SBOM Model'].notnull()
df4['aging']=df4.iloc[:,1:-3].ge(df4.iloc [:,-1],axis=0).sum(axis=1)
# For the rows where 'SBOM Model' is blank, set 'aging' to NaN
df4.loc[~mask, ['aging', 'Allocated']] = None
```

## Export Output


```python
# set the Output file path
os.chdir('D:/Advance_Data_Sets/BSS_EI_Output/TNL_Output')
with pd.ExcelWriter('05_Maximum_User_Number_Audit.xlsx',date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    # Audit
    df4.to_excel(writer,sheet_name="Maximum_User_Number",index=False)
```
