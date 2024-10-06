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
# Set Input File PATH
folder_path = 'D:/Advance_Data_Sets/GUL/GUL_Output'
os.chdir(folder_path)
```

## Import UMTS Mapping File


```python
df = pd.read_csv('02_UMTS_Cell_Mapping_wrt_Band.csv',usecols=['UMTS_Site_ID','PTML-U2100','PTML-U900','Project_Key'])
# fill missing values
df['PTML-U2100'] = df['PTML-U2100'].fillna('No Colocated U2100 Cell')
df['PTML-U900'] =  df['PTML-U900'].fillna('No Colocated U900 Cell')
```

## Import LTE Mapping File


```python
df0 = pd.read_csv('06_LTE_Cell_Mapping_wrt_Band.csv',\
                  usecols=['LTE_Site_ID','PTML-L-1800','PTML-L-2100','PTML-L-900','Project_Key'])
# fill missing values
df0['PTML-L-1800'] = df0['PTML-L-1800'].fillna('No Colocated L-1800 Cell')
df0['PTML-L-2100'] = df0['PTML-L-2100'].fillna('No Colocated L-2100 Cell')
df0['PTML-L-900']  = df0['PTML-L-900'].fillna('No Colocated L-900 Cell')
```

## Import Project Key


```python
df1 = pd.read_csv('08_Project_Key.csv')
```

## Merge Project Key data Frame with UMTS and LTE Mapping File


```python
df2 = pd.merge(df1,df,how='left',on='Project_Key').merge(df0,how='left',on='Project_Key')
```

# Fill Missing Values


```python
df2[['UMTS_Site_ID', 'PTML-U2100', 'PTML-U900']] = df2[['UMTS_Site_ID', 'PTML-U2100', 'PTML-U900']].fillna('UMTS Tech not deployed')
df2[['LTE_Site_ID','PTML-L-1800','PTML-L-2100','PTML-L-900']] = df2[['LTE_Site_ID','PTML-L-1800','PTML-L-2100','PTML-L-900']].fillna('LTE Tech not deployed')
```

## Extect Site_Key


```python
# Extect Site_Key
df2['Site_Key'] = df2['Project_Key'].apply(lambda x: x.split(';')[0]).astype(str)
# insert the column at specific lock and del from the last
df2.insert(0, 'Site_Key', df2.pop('Site_Key'))
```

## Drop unwanted coulumn


```python
df2 = df2.drop('Project_Key', axis=1)
```

## Rename all the Columns with Pre-Fix


```python
# re-name All the Columns
df2 = df2.add_prefix('NE Information;')
```

## Export Output


```python
folder_path = 'D:/Advance_Data_Sets/GUL/GUL_Output'
os.chdir(folder_path)
df2.to_csv('09_Project_Cell_Mapping.csv',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
