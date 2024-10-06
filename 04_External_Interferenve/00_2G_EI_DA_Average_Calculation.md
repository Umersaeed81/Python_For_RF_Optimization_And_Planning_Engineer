#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# External Interference KPIs (2G KPIs DA)

## Delete Files From the Folder, If Folder Not Found Created the Folder


```python
import os

# Folder path
folder_path = r"D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files_EI"

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

## Import required Libraries


```python
import os
import shutil
import zipfile
import numpy as np
import pandas as pd
from glob import glob
from datetime import datetime, timedelta
import warnings
warnings.simplefilter("ignore")
```

## Filter Required Files


```python
# List all files in the path
file_list = glob('D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/GSM/*.zip')

# Calculate the date for three days ago
required_date = datetime.now() - timedelta(days=10)
required_date_str = required_date.strftime("%d%m%Y")

# Filter files that match the required date or the two days before
filtered_file_list = [file for file in file_list if any((required_date + timedelta(days=i)).strftime("%d%m%Y") in file for i in range(10))]

if filtered_file_list:
    print("Filtered File List:", filtered_file_list)
else:
    print(f"No file found with the required dates in the last three days.")
```

## Import and Concat GSM DA KPIs


```python
df=pd.concat((pd.read_csv(file,skiprows=range(6),\
    skipfooter=1,engine='python',na_values=['NIL','/0'],\
    parse_dates=["Date"],\
    usecols=['Date','GBSC','Cell CI',\
    'Interference Band Proportion (4~5)(%)']) for file in filtered_file_list))
```

## Data Pre-Processing


```python
df0 = df.dropna().reset_index(drop=True)
```

## Calculate Average IOI (DA)


```python
df1=df0.groupby(['GBSC','Cell CI']).\
    apply(lambda x: np.average(x['Interference Band Proportion (4~5)(%)']))\
    .reset_index(name='2G_DA_IoI_Average')
```

## Export Output


```python
# Set Output File Path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files_EI'
os.chdir(folder_path)

with pd.ExcelWriter('00_External_Interference_GSM_DA.xlsx',date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    # GSM KPIs Values
    df1.to_excel(writer,sheet_name="GSM_IOI_DA",index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
