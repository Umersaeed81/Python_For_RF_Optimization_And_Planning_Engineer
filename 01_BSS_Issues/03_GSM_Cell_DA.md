#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

## BSS Issues (GSM Cell DA)

## Import Libraries


```python
# Import Libraries
import os
import zipfile
import numpy as np
import pandas as pd
from glob import glob
from datetime import datetime, timedelta
import warnings
warnings.simplefilter("ignore")
```

## GSM DA KPIs


```python
# Set Path for GSM Cell DA KPIs
path_cell_da = 'D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/GSM'
os.chdir(path_cell_da)
```

## Filter Required Files


```python
# List all files in the path
file_list = glob('*.zip')

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

    Filtered File List: ['GSM_Cell_DA_01102024.zip', 'GSM_Cell_DA_02102024.zip', 'GSM_Cell_DA_03102024.zip', 'GSM_Cell_DA_27092024.zip', 'GSM_Cell_DA_28092024.zip', 'GSM_Cell_DA_29092024.zip', 'GSM_Cell_DA_30092024.zip']
    

## Import Input Data Set


```python
# List zip file in a given folder
#df = sorted(glob('*.zip'))
# concat all the Cluster DA Files
df0=pd.concat((pd.read_csv(file,skiprows=range(6),\
    skipfooter=1,engine='python',
    usecols=['Date','Cell CI','GBSC','TCH Availability Rate(%)','Site Name',\
            'CM333:Call Drops due to Abis Terrestrial Link Failure (Traffic Channel)',\
            'CM334:Call Drops due to Equipment Failure (Traffic Channel)','Cell Name',\
            'A3129B:Failed Assignments (First Assignment, Terrestrial Resource Request Failed)'],\
            parse_dates=["Date"],na_values=['NIL','/0'],\
            dtype = {"Cell CI" : "str"}) for file in filtered_file_list)).reset_index(drop=True)
```

## Calculate BSS Drips


```python
# Calculate Cell DA KPIs
df0['BSS Drops']= df0['CM333:Call Drops due to Abis Terrestrial Link Failure (Traffic Channel)']\
                    +df0['CM334:Call Drops due to Equipment Failure (Traffic Channel)']
```

## Filter Required Columns


```python
# Filter Required Columns
df1 = df0[['Date','GBSC','Site Name','Cell CI','Cell Name',\
           'BSS Drops','A3129B:Failed Assignments (First Assignment, Terrestrial Resource Request Failed)',\
           'TCH Availability Rate(%)']]
```

## Re-Shape Data Set (Pivot_table)


```python
# re-share(pivot_table) (DA Level)
df2=df1.pivot_table\
    (index=["GBSC",'Site Name','Cell Name','Cell CI'],\
    columns="Date",values=['BSS Drops','A3129B:Failed Assignments (First Assignment, Terrestrial Resource Request Failed)',\
           'TCH Availability Rate(%)']).reset_index()
```

## Conditional Count


```python
# Count of Days TCH Availability Rate <100 (DA Level)
df2["Days (TCH Availability Rate)<100"] = (df2.iloc[:, 18:].lt(100)).sum(axis=1)
```

## Export Output


```python
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files'
os.chdir(folder_path)
```


```python
# Write excel file with default behaviour.
with pd.ExcelWriter("00_GSM_BSS_Issus.xlsx",date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    # GSM BSS Issues
    df2.to_excel(writer, sheet_name='GSM_BSS_Issus')
```


```python
# re-set all the variable from the RAM
%reset -f
```
