#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

## Import Libraries


```python
# Import Libraries
import os
import zipfile
import numpy as np
import pandas as pd
from glob import glob
from datetime import datetime, timedelta
import warnings
warnings.simplefilter("ignore")
```

## Input File Path


```python
# Set Path for GSM Cell DA KPIs
path_cell_hourly = 'D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files'
os.chdir(path_cell_hourly)
```

## UMTS Files Import and Processing


```python
# Down Time DA in Pivot table
df = pd.read_excel('02_UMTS_BSS_Issus.xlsx',dtype={'Cell ID':str})
# Number of intervals Down Time in 24hrs 
df0 = pd.read_excel('03_UMTS_BSS_Issus_Intervals.xlsx',sheet_name='UMTS_Ava_24hrs Interval Count',dtype={'Cell ID':str})
# Number of intervals Down Time in working hours 
df1 = pd.read_excel('03_UMTS_BSS_Issus_Intervals.xlsx',sheet_name='UMTS_Ava_9-21hrs Interval Count',dtype={'Cell ID':str})
# merge Data Set in Required Format
df2= pd.merge(df, df0[['RNC','Cell ID','Total Intervals (Down Time)>0 between 0:00-23:00']], on=['RNC','Cell ID'],how='left').merge(df1[['RNC','Cell ID','Total Intervals (Down Time)>0 between 9:00-21:00']], on=['RNC','Cell ID'],how='left')
```

## LTE Files Import and Processing


```python
# Down Time DA in Pivot table
df3 = pd.read_excel('04_LTE_BSS_Issus.xlsx',dtype={'LocalCell Id':str})
# Number of intervals Down Time in 24hrs 
df4 = pd.read_excel('05_LTE_BSS_Issus_Intervals.xlsx',dtype={'LocalCell Id':str},sheet_name='LTE_Ava_24hrs Interval Count')
# Number of intervals Down Time in working hours 
df5 = pd.read_excel('05_LTE_BSS_Issus_Intervals.xlsx',dtype={'LocalCell Id':str},sheet_name='LTE_Ava_9-21hrs Interval Count')
# merge Data Set in Required Format
df6= pd.merge(df3, df4[['eNodeB Name','Cell Name','LocalCell Id','eNodeB Function Name','Total Intervals (Down Time)>0 between 0:00-23:00']], on=['eNodeB Name','Cell Name','LocalCell Id','eNodeB Function Name'],how='left').merge(df5[['eNodeB Name','Cell Name','LocalCell Id','eNodeB Function Name','Total Intervals (Down Time)>0 between 9:00-21:00']], on=['eNodeB Name','Cell Name','LocalCell Id','eNodeB Function Name'],how='left')
```

## Export Output File


```python
# set path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files'
os.chdir(folder_path)

# Write excel file with default behaviour.
with pd.ExcelWriter("02_UMTS_LTE_BSS_Issus_Intervals.xlsx",date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    # UMTS BSS Issues
    df2.to_excel(writer, sheet_name='UMTS_BSS_Issus',index=False)
    df0.to_excel(writer, sheet_name='UMTS_Ava_24hrs Interval Count',index=False)
    df1.to_excel(writer, sheet_name='UMTS_Ava_9-21hrs Interval Count',index=False)
    # LTE BSS Issues
    df6.to_excel(writer, sheet_name='LTE_BSS_Issus',index=False)
    df4.to_excel(writer, sheet_name='LTE_Ava_24hrs Interval Count',index=False)
    df5.to_excel(writer, sheet_name='LTE_Ava_9-21hrs Interval Count',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```

tt
