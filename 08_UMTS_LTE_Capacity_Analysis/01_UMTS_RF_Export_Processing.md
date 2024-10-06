#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# UMTS RF Export (Cell File) Processing

## Import Required Libraries


```python
# Import Libraries
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

## Unzip Files


```python
# Set RF Export File Path
folder_path = 'D:/Advance_Data_Sets/RF_Export/UMTS/00_Cell_Export'
os.chdir(folder_path)

# Unzip Files 
for file in os.listdir(folder_path):   # get the list of files
    if zipfile.is_zipfile(file): # if it is a zipfile, extract it
        with zipfile.ZipFile(file) as zip_ref: # treat the file as a zip
            folder_name = os.path.splitext(file)[0]  # get the name of the folder (remove '.zip' extension)
            os.makedirs(folder_name, exist_ok=True)  # create the folder if it doesn't exist
            zip_ref.extractall(folder_name)  # extract the contents into the folder
```

## Filter CELL File


```python
df = glob('D:/Advance_Data_Sets/RF_Export/UMTS/00_Cell_Export/**/CELL.txt',recursive=True)
```

## Import and Concat Cell File


```python
# import and concat Cell File
df0=pd.concat((pd.read_csv(file,header=1,\
                 engine='python',encoding='unicode_escape',\
                 usecols=['BSC Name','Cell ID','Band Indicator',\
                          'Validation indication','Uplink UARFCN','Downlink UARFCN',\
                         'NodeB Name','Cell Name'],\
                 dtype={'Cell ID':str}) for file in df)).reset_index(drop=True)
```

## Replace multiple spaces with single space in 'eNodeB Name' & 'CELLNAME' column


```python
# Replace multiple spaces with single space in 'Cell Name' column
df0['Cell Name'] = df0['Cell Name'].str.replace(r'\s+', ' ', regex=True)


# Replace multiple spaces with single space in 'NodeB Name' column
df0['NodeB Name'] = df0['NodeB Name'].str.replace(r'\s+', ' ', regex=True)
```

## Replace Cell Name (atleast the audit once in a mounth)


```python
df0['Cell Name'] = df0['Cell Name'].replace({
    '3G-S-2654_Kashki Hamza Zai (USF) (S-7424)': '3G-S-2654-3_Kashki Hamza Zai (USF) (S-7424)',
    '3G-N-4243-K_Kot Azmath (Gold) Bannu (N-4757)': '3G-N-4243-4_Kot Azmath (Gold) Bannu (N-4757)'
})
```

## Generate UMTS Site/Cell IDs


```python
# Generate UMTS Site ID
df0['UMTS_Site_ID'] = df0['NodeB Name'].str.split('_').str[0]

# Generate UMTS Cell ID
df0['UMTS_Cell_ID'] = df0['Cell Name'].str.split('_').str[0]
```

## Band Identification


```python
# Band Identification
df0['Band'] = np.where(
            (df0['Downlink UARFCN'].eq(3012)),
            'PTML-U900', 
            np.where(
                    (df0['Downlink UARFCN'].eq(10637)), 
                    'PTML-U2100', 
                np.where(
                    (df0['Downlink UARFCN'].eq(10612)), 
                    'Telenor-U2100', 
                     'Other-Band')))
```

## Generate Key


```python
# Extract data after the last hyphen and add it as a new column
df0['UMTS_Key1'] = df0['UMTS_Site_ID']+";"+(df0['UMTS_Cell_ID'].apply(lambda x: x.split('-')[-1])).astype(str)
# Key : Key1 + Band
df0['UMTS_Key2'] = df0['UMTS_Key1']+";"+df0['Band']
```

## Identify Cell Names that are duplicated


```python
#Identify Cell Names that are duplicated
df1 = df0[df0.duplicated(subset='Cell Name', keep=False)].reset_index(drop=True)
```

## Filter ACTIVATED Cells From Duplicate Data Frame 


```python
df2 = df1[df1['Validation indication'].eq('ACTIVATED')]
```

## Filter out the duplicate Cell Names entirely


```python
# Filter out the duplicate Cell Names entirely
df3 = df0[~df0['Cell Name'].isin(df1['Cell Name'])].reset_index(drop=True)
```

## Append df2 and df3


```python
# Append df2 and df3
df4 = pd.concat([df2, df3], ignore_index=True)
```

## Delete Subfolders


```python
# List all items in the directory
items = os.listdir(folder_path)

# Iterate through the items and delete subfolders
for item in items:
    item_path = os.path.join(folder_path, item)
    if os.path.isdir(item_path):
        shutil.rmtree(item_path)
```

## Export RF Export Output


```python
folder_path = 'D:/Advance_Data_Sets/GUL/GUL_Output'
os.chdir(folder_path)
df4.to_csv('00_UMTS_RF_Export.csv',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
