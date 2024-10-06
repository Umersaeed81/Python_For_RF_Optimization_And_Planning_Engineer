#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

## BSS Issues (LTE Cell DA)

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

## LTE DA KPIs


```python
# File Path
path = 'D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/LTE'
os.chdir(path)
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

    Filtered File List: ['LTE_Cell_DA_01102024.zip', 'LTE_Cell_DA_02102024.zip', 'LTE_Cell_DA_03102024.zip', 'LTE_Cell_DA_27092024.zip', 'LTE_Cell_DA_28092024.zip', 'LTE_Cell_DA_29092024.zip', 'LTE_Cell_DA_30092024.zip']
    

## Import Input Data Set


```python
# List Zip Files in a Folder
#df = sorted(glob('*.zip'))

# concat all the Cell DA Files
df0=pd.concat((pd.read_csv(file,skiprows=range(6),\
    skipfooter=1,engine='python',\
    parse_dates=["Date"],na_values=['NIL','/0'],\
    usecols=['Date', 'eNodeB Name','Cell Name','LocalCell Id','eNodeB Function Name',
       'Radio Network Unavailability Rate_Cell'],\
         dtype={'LocalCell Id':int}) for file in filtered_file_list)).reset_index(drop=True)
```

## Re-share (pivot_table)


```python
# re-share (pivot_table) of Radio Network Unavailability Rate_Cell
df1 = pd.pivot_table(df0, \
          values='Radio Network Unavailability Rate_Cell', \
          index=['eNodeB Name','Cell Name','LocalCell Id','eNodeB Function Name'],columns='Date').reset_index()
```

## Conditional Count


```python
#Count of Days Cell Unavailability duration (Down Time) >0 (DA Level)
df1["Days (Down Time)>0"] = (df1.iloc[:, 4:].gt(0)).sum(axis=1)
```

## Export Output


```python
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files'
os.chdir(folder_path)
# Write excel file with default behaviour.
with pd.ExcelWriter("04_LTE_BSS_Issus.xlsx",date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    # LTE BSS Issues
    df1.to_excel(writer, sheet_name='LTE_BSS_Issus',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```


```python

```
