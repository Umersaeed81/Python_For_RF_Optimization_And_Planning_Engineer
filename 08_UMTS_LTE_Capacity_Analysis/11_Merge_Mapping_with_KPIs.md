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

## Import Project Cell Mapping File


```python
df = pd.read_csv('09_Project_Cell_Mapping.csv')
```

## Import UMTS Process KPIs


```python
# For U-900
df0 = pd.read_csv('03_UMTS_Cell_Process_KPIs.csv',\
                usecols=lambda column: column != 'Cell Name').\
                add_prefix('PTML-U900;').rename(columns={'PTML-U900;UMTS_Cell_ID': 'NE Information;PTML-U900'})

# For U-2100
df1 = pd.read_csv('03_UMTS_Cell_Process_KPIs.csv',\
                usecols=lambda column: column != 'Cell Name').\
                add_prefix('PTML-U2100;').rename(columns={'PTML-U2100;UMTS_Cell_ID': 'NE Information;PTML-U2100'})
```

## Import LTE Process KPIs


```python
# For L-900
df2 = pd.read_csv('07_LTE_Cell_Process_KPIs.csv',\
                usecols=lambda column: column != 'Cell Name').\
                add_prefix('PTML-L-900;').rename(columns={'PTML-L-900;LTE_Cell_ID': 'NE Information;PTML-L-900'})


# For L-1800
df3 = pd.read_csv('07_LTE_Cell_Process_KPIs.csv',\
                usecols=lambda column: column != 'Cell Name').\
                add_prefix('PTML-L-1800;').rename(columns={'PTML-L-1800;LTE_Cell_ID': 'NE Information;PTML-L-1800'})


# For L-2100
df4 = pd.read_csv('07_LTE_Cell_Process_KPIs.csv',\
                usecols=lambda column: column != 'Cell Name').\
                add_prefix('PTML-L-2100;').rename(columns={'PTML-L-2100;LTE_Cell_ID': 'NE Information;PTML-L-2100'})
```

## Step-1: Merge Project Cell Mapping File with U-2100 KPIs (df1)


```python
df5 = pd.merge(df,df1,how='left',on=['NE Information;PTML-U2100'])
```

## Step-2: Merge Step-1(df5) with U-900 KPIs (df0)


```python
df6 = pd.merge(df5,df0,how='left',on=['NE Information;PTML-U900'])
```

## Step-3 : Merge Step-2(df6) with L-1800 KPIs (df3)


```python
df7 = pd.merge(df6,df3,how='left',on=['NE Information;PTML-L-1800'])
```

## Step-4: Merge Step-3(df7) with L-2100 KPIs(df4)


```python
df8 = pd.merge(df7,df4,how='left',on=['NE Information;PTML-L-2100'])
```

## Step-5: Merge Step-4 (df8) with L-900 KPIs(df2)


```python
df9 = pd.merge(df8,df2,how='left',on=['NE Information;PTML-L-900'])
```

## Convert the Data Frame to Multindex Header


```python
# Convert the Data Frame to Multindex Header
idx = df9.columns.str.split(';', expand=True)
df9.columns=idx
```

## Site Mapping


```python
# Set Input File Path
folder_path = 'D:/Advance_Data_Sets/GUL/Mapping'
os.chdir(folder_path)
df10 = pd.read_excel('Mapping.xlsx',dtype={'GSM_Site_ID': str,'UMTS_Site_ID': str , 'LTE_Site_ID':str , 'Site_Key':str})
```

## Output


```python
folder_path = 'D:/Advance_Data_Sets/GUL/GUL_Output'
os.chdir(folder_path)
```


```python
with pd.ExcelWriter('10_UMTS_LTE_Capacity_Analysis_{}.xlsx'.format((pd.Timestamp.today()).strftime('%d%m%Y'))) as writer:
        df10.to_excel(writer, sheet_name="Site_Mapping",engine='openpyxl', na_rep='-',index=False)
        df9.to_excel(writer, sheet_name="GUL Utilization",engine='openpyxl', na_rep='-')
```


```python
# re-set all the variable from the RAM
%reset -f
```
