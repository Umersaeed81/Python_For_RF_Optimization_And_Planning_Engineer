#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# GSM RF Untilization (TNX Type Identification)

## Import requied Libraries


```python
import os
import re
import glob
import zipfile
import logging
import pandas as pd
from glob import glob
from datetime import datetime

import warnings
warnings.simplefilter("ignore")
```

## Import and Concat GSM BTS DA KPIs


```python
# set file path
working_directory = 'D:/Advance_Data_Sets/TXN_DataSets/GSM_BTS_DA'
os.chdir(working_directory)

# List all files in the path
file_list = glob('*.zip')
logging.info(f"List of files: {file_list}")

# Extract dates and convert them to datetime objects
date_pattern = re.compile(r'_(\d{8})\.zip$')
dates = []

for file in file_list:
    match = date_pattern.search(file)
    if match:
        date_str = match.group(1)
        date_obj = datetime.strptime(date_str, '%d%m%Y')
        dates.append((date_obj, file))
        logging.info(f"Extracted date from {file}: {date_obj}")

# Find the file with the maximum date
if dates:
    max_date_file = max(dates, key=lambda x: x[0])[1]
    max_date_file_list = [max_date_file]
    logging.info(f'The file with the latest date is: {max_date_file_list}')
else:
    max_date_file_list = []
    logging.warning('No dates extracted from the files.')

print(max_date_file_list)

```


```python
# concat all the BTS DA Files
df0=pd.concat((pd.read_csv(file,skiprows=range(6),\
            skipfooter=1,engine='python',\
            usecols=['Date','GBSC','Site Name','R2741:Number of Available Flex Timeslots'],
            parse_dates=["Date"],na_values=['NIL','/0']) for file in max_date_file_list))\
            .sort_values('Date').set_index(['Date']).\
            last('1D').reset_index().rename(columns={'R2741:Number of Available Flex Timeslots':'R2741'})
```

## Resource_Report


```python
# set file path
working_directory = 'D:/Advance_Data_Sets/Resource_RAN_Report/GSM/Resource_Report'
os.chdir(working_directory)
```

## Unzip Files (Zip File)


```python
for file in os.listdir(working_directory):   # get the list of files
    if zipfile.is_zipfile(file): # if it is a zipfile, extract it
        with zipfile.ZipFile(file) as zip_ref: # treat the file as a zip
            folder_name = os.path.splitext(file)[0]  # get the name of the folder (remove '.zip' extension)
            os.makedirs(folder_name, exist_ok=True)  # create the folder if it doesn't exist
            zip_ref.extractall(folder_name)  # extract the contents into the folder
```

## Unzip Files (rat File)


```python
import patoolib
# Get a list of all .rar files in the directory
rar_files = [file for file in os.listdir() if file.endswith(".rar")]

# Extract all .rar files
for rar_file in rar_files:
    patoolib.extract_archive(rar_file, outdir=".")
```

## Import and Concat Resource Report


```python
# Get Cell File Path from all the Sub folders
df1 = glob('D:/Advance_Data_Sets/Resource_RAN_Report/GSM/Resource_Report/**/BTS Timeslot Resource Summary.csv', recursive=True) 
# import and concat Resource Report File
df2=pd.concat((pd.read_csv(file,header=0,
                engine='python',encoding='unicode_escape',\
                usecols=['BSC Name','BTS Name','Service Type'],
                dtype={'BSC Name':str,'BTS Name':str,'Service Type':str}) for file in df1)).\
                drop_duplicates().reset_index(drop=True).\
                rename(columns={'BSC Name':'GBSC','BTS Name':'Site Name'})
```

## Delte Sub Folder


```python
import shutil
# Python Script for Clearing Subfolders in a Specific path
for subfolder in os.listdir(working_directory):
    subfolder_path = os.path.join(working_directory, subfolder)
    if os.path.isdir(subfolder_path):
        shutil.rmtree(subfolder_path)
```

## Merge KPIs and Resource Report


```python
df3 = pd.merge(df0,df2,how='left',on=['GBSC','Site Name']).fillna('-')
```

## TXN Type Identification


```python
import numpy as np
df3['TNX_Type'] = np.where(
    (df3['R2741'].eq(0)) & df3['Service Type'].eq('IP'),
    'IP',
    np.where(
        (df3['R2741'].eq(0)) & df3['Service Type'].eq('TDM'),
        'Fix',
        np.where(
            (df3['R2741'].ne(0)) & df3['Service Type'].eq('TDM'),
            'Flex',
            '-')))
```

## Select Required Columns


```python
df4 = df3[['GBSC','Site Name','TNX_Type']]
```

## Output


```python
# Set Output Path
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(folder_path)
# Export
df4.to_csv('06_TXN_Type.csv',index=False)
```

## Re-Set Variables


```python
#re-set all the variable from the RAM
%reset -f
```
