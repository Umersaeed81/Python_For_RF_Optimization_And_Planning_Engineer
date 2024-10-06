#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

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

## Set RF Export File Path


```python
# Set RF Export File Path
folder_path = 'D:/Advance_Data_Sets/RF_Export/UMTS/00_Cell_Export'
os.chdir(folder_path)
```

## Unzip UMTS RF Export


```python
for file in os.listdir(folder_path):   # get the list of files
    if zipfile.is_zipfile(file): # if it is a zipfile, extract it
        with zipfile.ZipFile(file) as zip_ref: # treat the file as a zip
            folder_name = os.path.splitext(file)[0]  # get the name of the folder (remove '.zip' extension)
            os.makedirs(folder_name, exist_ok=True)  # create the folder if it doesn't exist
            zip_ref.extractall(folder_name)  # extract the contents into the folder
```

## Import and concat RF Export Cell File


```python
# Filter requird files
df = glob('D:/Advance_Data_Sets/RF_Export/UMTS/00_Cell_Export/**/CELL.txt', recursive=True)
# import and concont required files
df0=pd.concat((pd.read_csv(file,header=1,\
                skipfooter=1,engine='python',\
                usecols=['BSC Name','Cell ID','Band Indicator',\
                'Validation indication'],encoding='unicode_escape',dtype={'Cell ID':str}) for file in df)).\
                rename(columns = {"BSC Name":"RNC"})
```

## Set Input File Path


```python
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files_EI'
os.chdir(folder_path)
```

## Import Average RTWP Value Sheet 


```python
df1 = pd.read_excel('01_External_Interference_UMTS_DA.xlsx',dtype={'Cell ID':str})
```

## Merge Average KPI Value with RF Export


```python
df2 = pd.merge(df0,df1,on=['RNC','Cell ID'],how='left')
```

## Import Interval Sheet


```python
df3 = pd.read_excel('04_External_Interference_UMTS_Intervals.xlsx',sheet_name='UMTS_RTWP_Intervals_-95',dtype={'Cell ID':str})
df4 = pd.read_excel('04_External_Interference_UMTS_Intervals.xlsx',sheet_name='UMTS_RTWP_Intervals_-98',dtype={'Cell ID':str})
```

## Merge df2 with Interval Sheets


```python
df5 = pd.merge(df2,df3[['RNC','Cell ID','Total_Interval_Impacted_RTWP>=-95']],on=['RNC','Cell ID'],how='left').merge(df4[['RNC','Cell ID','Total_Interval_Impacted_RTWP>=-98']],on=['RNC','Cell ID'],how='left')
```

## Import Max RTWP Sheet


```python
df6 = pd.read_excel('07_External_Interference_UMTS_Max.xlsx',dtype={'Cell ID':str})
```

## Merge df5 with df6


```python
df7 = pd.merge(df5,df6,on=['RNC','Cell ID'],how='left')
```

## Concat two columns at specific location


```python
df7.insert(0, 'RNC;Cell ID', df7['RNC'] + ';' + df7['Cell ID'].astype(str))
```

## Insert Blank Columns


```python
df7[['RACaseSerialNo', 'Region', 'City', 'Category', 'Status(Open/Close)']] = ''
```

## Export Output


```python
with pd.ExcelWriter('10_External_Interference_UMTS_Final.xlsx',date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    # UMTS KPIs Values
    for value in df7['Band Indicator'].unique():
        df7[df7['Band Indicator'] == value].\
        to_excel(writer, index=False, sheet_name=f'{value}')
    # Export U900 Intervals
    df3.to_excel(writer, index=False, sheet_name='UMTS_U900_RTWP_Intervals')
    # Export U2100 Intervals
    df4.to_excel(writer, index=False, sheet_name='UMTS_U2100_RTWP_Intervals')
```


```python
# re-set all the variable from the RAM
%reset -f
```
