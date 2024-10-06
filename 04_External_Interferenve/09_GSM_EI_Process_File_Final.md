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

## Set Input File Path


```python
# Set Path for GSM Cell DA KPIs
file_path = 'D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files_EI'
os.chdir(file_path)
```

## GSM IOI DA KPIs


```python
df = pd.read_excel('00_External_Interference_GSM_DA.xlsx',dtype={'Cell CI':str})
```

## GSM Intervals Count


```python
df0 = pd.read_excel('03_External_Interference_GSM_Intervals.xlsx',dtype={'Cell CI':str})
```

## GSM Max Values


```python
df1 = pd.read_excel('06_External_Interference_GSM_Max.xlsx',dtype={'Cell CI':str})
```

## Merge Data Frame


```python
df2 = pd.merge(df,df0[['GBSC','Cell CI','Total_Intervals_Impacted']],how='left',on=['GBSC','Cell CI']).\
                                                                merge(df1,how='left',on=['GBSC','Cell CI'])
```

## Concat two columns at specific location


```python
df2.insert(0, 'GBSC;Cell CI', df2['GBSC'] + ';' + df2['Cell CI'].astype(str))
```

## Insert Blank Columns


```python
df2[['RACaseSerialNo', 'Region', 'City', 'Category', 'Status(Open/Close)']] = ''
```

## Export Output Data Frame


```python
# Set Output File Path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files_EI'
os.chdir(folder_path)

with pd.ExcelWriter('09_External_Interference_GSM_Final.xlsx',date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    # GSM IOI (DA, Impacted Intervals, Max IOI, )
    df2.to_excel(writer,sheet_name="GSM_IOI_Huawei",index=False)
    # GSM Impacted Intervals per day count
    df0.to_excel(writer,sheet_name="GSM_IOI_Intervals",index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
