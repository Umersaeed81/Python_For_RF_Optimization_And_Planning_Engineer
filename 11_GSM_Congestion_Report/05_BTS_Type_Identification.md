#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# GSM RF Untilization (BTS Type Identification)

## Import requied Libraries


```python
import os
import glob
from glob import glob
import zipfile
import pandas as pd
import warnings
warnings.simplefilter("ignore")
```

## Unzip 2G RF Export (Cell Level)


```python
# set file path
working_directory = 'D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export'
os.chdir(working_directory)
# unzip Export
for file in os.listdir(working_directory):   # get the list of files
    if zipfile.is_zipfile(file): # if it is a zipfile, extract it
        with zipfile.ZipFile(file) as zip_ref: # treat the file as a zip
            folder_name = os.path.splitext(file)[0]  # get the name of the folder (remove '.zip' extension)
            os.makedirs(folder_name, exist_ok=True)  # create the folder if it doesn't exist
            zip_ref.extractall(folder_name)  # extract the contents into the folder
```

## Import and Concat BTSCELLPATCHPARA File


```python
# Get Cell File Path from all the Sub folders
df = glob('D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export/**/BTSCELLPATCHPARA.txt', recursive=True) 

# import and concat Resource Report File
df0=pd.concat((pd.read_csv(file,header=0,skiprows=[1],
                engine='python',encoding='unicode_escape',\
                usecols=['BSCName','BTSID','CELLID'],
                dtype={'BSCName':str,'BTSID':str,'CELLID':str}) for file in df)).\
                drop_duplicates().reset_index(drop=True)
```

## Import and Concat GCELL File


```python
# Get Cell File Path from all the Sub folders
df1 = glob('D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export/**/GCELL.txt', recursive=True) 

# import and concat Resource Report File
df2=pd.concat((pd.read_csv(file,header=0,skiprows=[1],
                engine='python',encoding='unicode_escape',\
                usecols=['BSCName','BTSNAME','CELLNAME','CELLID'],
                dtype={'BSCName':str,'BTSNAME':str,'CELLNAME':str,'CELLID':str}) for file in df1)).\
                drop_duplicates().reset_index(drop=True)
```

## Merge GCELL with BTSCELLPATCHPARA (Ref-1)


```python
df3 = pd.merge(df2,df0,how='left',on=['BSCName','CELLID']).fillna('-')
```

## RAN Report (Unzip)


```python
# set file path
working_directory = 'D:/Advance_Data_Sets/Resource_RAN_Report/GSM/RAN_Report'
os.chdir(working_directory)
# unzip Export
for file in os.listdir(working_directory):   # get the list of files
    if zipfile.is_zipfile(file): # if it is a zipfile, extract it
        with zipfile.ZipFile(file) as zip_ref: # treat the file as a zip
            folder_name = os.path.splitext(file)[0]  # get the name of the folder (remove '.zip' extension)
            os.makedirs(folder_name, exist_ok=True)  # create the folder if it doesn't exist
            zip_ref.extractall(folder_name)  # extract the contents into the folder
```

## RAN Report (import and concat)


```python
# Use glob to find all files with the name 'GSM_Site_Report*.csv' recursively
df4 = glob(os.path.join(working_directory, '**', 'GSM_Site_Report*.csv'), recursive=True)

# import and concat Resource Report File
df5=pd.concat((pd.read_csv(file,header=0,
                engine='python',encoding='unicode_escape',\
                usecols=['NE Name','Site Index','Site Type'],
                dtype={'NE Name':str,'Site Index':str,'Site Type':str}) for file in df4)).\
                drop_duplicates().reset_index(drop=True).\
                rename(columns={'NE Name':'BSCName','Site Index':'BTSID'})
```

## Merge Ref-1 with RAN Report


```python
df6 = pd.merge(df3,df5,how='left',on=['BSCName','BTSID']).fillna('-')
```


```python
# Set Output Path
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(folder_path)
# Export
df6.to_csv('05_BTS_Type.csv',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
