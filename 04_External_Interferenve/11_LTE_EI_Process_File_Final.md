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

## Import and Concatenate RF Export


```python
# Filter Required Files
df = sorted(glob('D:/Advance_Data_Sets/RF_Export/LTE/Cell/*.xlsx'))
# Concatenate 'Cell Sheet' from All the Excel File
df0 = pd.concat(pd.read_excel(f, sheet_name='CELL', engine="openpyxl", \
                header=1,usecols=['*eNodeB Name','*Local Cell ID','*Cell Name']) for f in df).\
                 rename(columns={"*Cell Name":"Cell Name"})   

# create a cell ID
df0['Cell ID'] = df0['*Local Cell ID'].astype(str) +""+\
                        df0['*eNodeB Name'].str.split('_').str[0].str.split('-').str[-1].astype(str)
```

## Set Input File Path


```python
# Set Path for GSM Cell DA KPIs
file_path = 'D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files_EI'
os.chdir(file_path)
```


```python
df1 = pd.read_excel('02_External_Interference_LTE_DA.xlsx')
df2 = pd.read_excel('05_External_Interference_LTE_Intervals.xlsx')
df3 = pd.read_excel('08_External_Interference_LTE_Max.xlsx')
```

## Meger df1 with df0


```python
# merge dateframe (df1 and df2)
df4 = pd.merge(df1,df0[['Cell Name','Cell ID']],on=['Cell Name'],how='left')
# Move columne at a specific Location
df4.insert(1, 'Cell ID', df4.pop('Cell ID'))
```

## Mege df4 with df2


```python
df5 = pd.merge(df4,df2[['Cell Name','Total_Intervals_Impacted']],on=['Cell Name'],how='left')
```

## Merge df5 with df3


```python
df6 = pd.merge(df5,df3,on=['Cell Name'],how='left')
```

## Insert Blank Columns


```python
df6[['RACaseSerialNo', 'Region', 'City', 'Category', 'Status(Open/Close)']] = ''
```

## Export Output


```python
with pd.ExcelWriter('11_External_Interference_LTE_Final.xlsx',date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    # LTE KPIs Values
    df6.to_excel(writer,sheet_name="LTE_UL_Interference_Huawei",index=False)
    # LTE Intervals
    df2.to_excel(writer,sheet_name="LTE_UL_Interference_Intervals",index=False)
```
