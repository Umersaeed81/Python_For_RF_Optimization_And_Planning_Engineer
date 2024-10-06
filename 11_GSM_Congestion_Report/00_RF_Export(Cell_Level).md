#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# GSM RF Untilization (Cell Level Paramets)

## Import requied Libraries


```python
import io
import os
import glob
import fnmatch
import zipfile
import pandas as pd
from itertools import permutations
```

## Set File Path


```python
# Set RF Export File Path
folder_path = 'D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export'
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

## GCELLCHMGAD


```python
# Get GCELLCHMGAD Path from all the Sub folders
df = glob.glob('D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export/**/GCELLCHMGAD.txt',\
                   recursive=True) 
# import and concat GCELLCHMGAD File
df0=pd.concat((pd.read_csv(file,header=0,skiprows=[1],
                engine='python',encoding='unicode_escape',\
                usecols=['BSCName','CELLNAME',
                'TCHBUSYTHRES','AMRTCHHPRIORLOAD',
                'TCHTRICBUSYOVERLAYTHR','TCHTRIBUSYUNDERLAYTHR',
                'LOWRXLEVOLFORBIDSWITCH']) for file in df)).\
                drop_duplicates().reset_index(drop=True)
```

## GCELLHOIUO


```python
# Get GCELLHOIUO File Path from all the Sub folders
df1 = glob.glob('D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export/**/GCELLHOIUO.txt',\
                recursive=True) 
# import and GCELLHOIUO Cell File
df2=pd.concat((pd.read_csv(file,header=0,skiprows=[1],
                engine='python',encoding='unicode_escape',\
                usecols=[
                'BSCName','CELLNAME','OPTILAYER',
                'HOALGOPERMLAY','ACCESSOPTILAY','OTOURECEIVETH','UTOORECTH']) for file in df1)).\
                drop_duplicates().reset_index(drop=True)
```

## GCELLVAMOS


```python
# Get GCELLVAMOS File Path from all the Sub folders
df3 = glob.glob('D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export/**/GCELLVAMOS.txt',\
                recursive=True) 
# import and GCELLVAMOS Cell File
df4=pd.concat((pd.read_csv(file,header=0,skiprows=[1],
                engine='python',encoding='unicode_escape',\
                usecols=['BSCName','CELLNAME','VAMOSSWITCH']) 
                                        for file in df3)).\
                drop_duplicates().reset_index(drop=True)
```

## GCELLCHMGBASIC


```python
# Get GCELLCHMGBASIC File Path from all the Sub folders
df5 = glob.glob('D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export/**/GCELLCHMGBASIC.txt', recursive=True) 
# import and concat GCELLCHMGBASIC Cell File
df6=pd.concat((pd.read_csv(file,header=0,skiprows=[1],
                engine='python',encoding='unicode_escape',\
                usecols=[
                'BSCName','CELLNAME','IDLESDTHRES','CELLMAXSD']) for file in df5)).\
                drop_duplicates().reset_index(drop=True)
```

## Merge RF Exports (GCELLCHMGAD,GCELLHOIUO,GCELLVAMOS,GCELLCHMGBASIC)


```python
df7 = df0.merge(df2,on=['BSCName','CELLNAME']).merge(df4,on=['BSCName','CELLNAME']).merge(df6,on=['BSCName','CELLNAME'])
```

## add_suffix in Column name


```python
# Add comma as suffix to column names
df7 = df7.add_prefix('Cell Level Parameters;')
```

## re-name column name


```python
# Rename the columns of the dataframe
df7 = df7.rename(columns={'Cell Level Parameters;BSCName': 'NE Information;BSCName',\
                   'Cell Level Parameters;CELLNAME':'NE Information;CELLNAME'})
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

## This script creates a folder named '2G_Utilization' in the specified path or, if it already exists, removes all the files within it.


```python
import shutil

def create_or_clear_folder(folder_path):
    if os.path.exists(folder_path):
        # Folder exists, clear its contents
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")
    else:
        # Folder doesn't exist, create it
        os.makedirs(folder_path)

# Define the folder path
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'

# Call the function to create or clear the folder
create_or_clear_folder(folder_path)
```

## Export RF Export (Cell Level Parameters)


```python
# Set Output Path
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(folder_path)
# Export
df7.to_csv('00_RF_Export_Cell_Para.csv',index=False)
```

## Re-Set Variables


```python
# re-set all the variable from the RAM
%reset -f
```
