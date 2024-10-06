#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

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

## Set Input File Path


```python
# Set Input File Path
folder_path = 'D:/Advance_Data_Sets/GUL/GUL_Output'
os.chdir(folder_path)
```

## Import UMTS RF Export Processing File


```python
df = pd.read_csv('00_UMTS_RF_Export.csv')
```

## Import Processed KPIs Cell ID


```python
df0 = pd.read_csv('01_UMTS_KPIs_Cell_List.csv',usecols=['UMTS_Cell_ID','Comments'])
```

## Merge Both Data Frame (in the case df and df0)


```python
df1 = pd.merge(df,df0,on=['UMTS_Cell_ID'],how='left')
```

## Fillna value in a Specific Column


```python
df1['Comments'] = df1['Comments'].fillna('No')
```

## Filter Conditioal Filter DataFrame


```python
df2 = df1[df1['Comments']=='Yes']
```

## Generate Key


```python
# calculate cumcount of 'UMTS_Key2'
df2['Count'] = df2.groupby('UMTS_Key2').cumcount() + 1
# key: 'UMTS_Key1' + cumcount
df2['UMTS_Key3'] = df2['UMTS_Key1']+";"+(df2['Count']).astype(str)
```

## UMTS Cell Mapping


```python
df3 = df2.pivot_table(index=['UMTS_Site_ID','UMTS_Key1','UMTS_Key3'],\
                    columns='Band',values='UMTS_Cell_ID',\
                    aggfunc=lambda x: ' '.join(str(v) for v in x)).reset_index()
```

## Import Site Mapping Sheet


```python
# Set Input File Path
folder_path = 'D:/Advance_Data_Sets/GUL/Mapping'
os.chdir(folder_path)
df4 = pd.read_excel('Mapping.xlsx',dtype={'GSM_Site_ID': str,'UMTS_Site_ID': str , 'LTE_Site_ID':str , 'Site_Key':str})
```

# Merge UMTS Cell Mapping DataFrame with Mapping Sheet


```python
df5 = pd.merge(df3,df4[['UMTS_Site_ID','Site_Key']],on='UMTS_Site_ID',how='left')
```

## Fill Missing Values


```python
# Fill missing values in 'Key' with values from 'UMTS_Site_ID'
df5['Site_Key'] = df5['Site_Key'].fillna(df5['UMTS_Site_ID'])
```

## Project Key


```python
df5['Project_Key'] = df5['Site_Key']+";"+(df5['UMTS_Key1'].apply(lambda x: x.split(';')[1])).astype(str)
```

## Export


```python
# Set Output Folder Path
folder_path = 'D:/Advance_Data_Sets/GUL/GUL_Output'
os.chdir(folder_path)
df5.to_csv('02_UMTS_Cell_Mapping_wrt_Band.csv',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
