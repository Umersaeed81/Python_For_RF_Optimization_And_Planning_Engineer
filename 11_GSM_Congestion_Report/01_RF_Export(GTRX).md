#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# GSM RF Untilization (GTRX Export)

## Import requied Libraries


```python
import io
import os
import glob
import fnmatch
import zipfile
import numpy as np
import pandas as pd
from itertools import permutations
```

## Set File Path


```python
# Set RF Export File Path
folder_path = 'D:/Advance_Data_Sets/RF_Export/GSM/03_GTRX_Export'
os.chdir(folder_path)
```

## Unzip File


```python
for file in os.listdir(folder_path):   # get the list of files
    if zipfile.is_zipfile(file): # if it is a zipfile, extract it
        with zipfile.ZipFile(file) as zip_ref: # treat the file as a zip
            folder_name = os.path.splitext(file)[0]  # get the name of the folder (remove '.zip' extension)
            os.makedirs(folder_name, exist_ok=True)  # create the folder if it doesn't exist
            zip_ref.extractall(folder_name)  # extract the contents into the folder
```

## GTRXCHAN


```python
# Get GCELL Path from all the Sub folders
df = glob.glob('D:/Advance_Data_Sets/RF_Export/GSM/03_GTRX_Export/**/GTRXCHAN.txt', recursive=True) 
# import and concat GCELL File
df0=pd.concat((pd.read_csv(file,header=0,skiprows=[1],
                engine='python',encoding='unicode_escape',\
                usecols=['BSCName','BTSNAME','CELLNAME','CHTYPE']) for file in df)).\
                reset_index(drop=True)
```

## Calculate Channel Type (Per Cell Level)


```python
df1=pd.crosstab([df0["BSCName"],df0["CELLNAME"]],\
                     df0['CHTYPE']).reset_index().fillna(0)
```

## add_suffix in Column name


```python
# Add comma as suffix to column names
df1 = df1.add_prefix('Channel Type;')
```

## Rename Header


```python
# Rename the columns of the dataframe
df1 = df1.rename(columns={'Channel Type;BSCName': 'NE Information;BSCName',\
                           'Channel Type;CELLNAME':'NE Information;CELLNAME'})
```

## Re-move Sub Folders


```python
# import libraries
import shutil

# Python Script for Clearing Subfolders in a Specific path
for subfolder in os.listdir(folder_path):
    subfolder_path = os.path.join(folder_path, subfolder)
    if os.path.isdir(subfolder_path):
        shutil.rmtree(subfolder_path)
```

## Export


```python
# Set Output Path
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(folder_path)
df1.to_csv('01_RF_Export_Ch_Type.csv',index=False)
```

## Re-Set Variables


```python
# re-set all the variable from the RAM
%reset -f
```
