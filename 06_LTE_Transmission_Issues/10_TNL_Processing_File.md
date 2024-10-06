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
# Required Libraries
import os
import numpy as np
import pandas as pd
from glob import glob
```

## Input File Path


```python
# set the input file path
os.chdir('D:/Advance_Data_Sets/BSS_EI_Output/TNL_Output')
```

## Import Data Sets


```python
# TNL Fail
df = pd.read_excel('00_LTE_TXN_Availability_Issues.xlsx',sheet_name='LTE_TNL_Fail')
# Abnormal TNL Fail
df0 = pd.read_excel('00_LTE_TXN_Availability_Issues.xlsx',sheet_name='LTE_AbnormRel_TNL_Fail')
# DA Availability
df1 = pd.read_excel('00_LTE_TXN_Availability_Issues.xlsx',sheet_name='LTE_Availability_Issues')
```


```python
# TNL Fail after ex. Low Availability Intervals
df2 = pd.read_excel('01_LTE_TXN_Availability_Issues.xlsx',sheet_name='LTE_Pure_TNL')

# Low Availability Intervals 24hrs
df3 = pd.read_excel('01_LTE_TXN_Availability_Issues.xlsx',sheet_name='LTE_Ava_24hrs(Interval Count)')

# Low Availability Intervals working hrs
df4 = pd.read_excel('01_LTE_TXN_Availability_Issues.xlsx',sheet_name='LTE_Ava_9-21hrs(Interval Count)')
```


```python
# Link Capacity
df5 = pd.read_excel('02_LTE_TXN_Availability_Issues.xlsx',sheet_name='LTE_TXN_Link_Capacity')
```

## Merge Required Data Frame


```python
df6 = df1.merge(df3[['eNodeB Name','Total Intervals (Down Time)>0 between 0:00-23:00']],how='left',on=['eNodeB Name']).\
        merge(df4[['eNodeB Name','Total Intervals (Down Time)>0 between 9:00-21:00']],how='left',on=['eNodeB Name'])
```


```python
# S1_HO_Attempts
df7 = pd.read_excel('00_LTE_TXN_Availability_Issues.xlsx',sheet_name='S1_HO_Attempts')
```


```python
# S1_HO_Attempts
df8 = pd.read_excel('03_LTE_TXN_Availability_Issues.xlsx',sheet_name='Flash_CSFB')
```


```python
# Base_Station_Transport_Data
df9 = pd.read_excel('04_LTE_TXN_Availability_Issues.xlsx',sheet_name='Base_Station_Transport_Data')
```


```python
# Max user File
df10 = pd.read_excel('05_Maximum_User_Number_Audit.xlsx',sheet_name='Maximum_User_Number')
```


```python
df11 = pd.read_excel('GLOBALPROCSWITCH.xlsx',sheet_name='LTE_eNodeB_SON=OFF')
```

## Export Output


```python
# Write excel file with default behaviour.
with pd.ExcelWriter("LTE_TXN_Availability_Issues.xlsx",date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    df.to_excel(writer, sheet_name='LTE_TNL_Fail',index=False)
    df0.to_excel(writer, sheet_name='LTE_AbnormRel_TNL_Fail',index=False)
    df2.to_excel(writer, sheet_name='LTE_Pure_TNL',index=False)
    df6.to_excel(writer, sheet_name='LTE_Availability_Issues',index=False)    
    df3.to_excel(writer, sheet_name='LTE_Ava_24hrs(Interval Count)',index=False)
    df4.to_excel(writer, sheet_name='LTE_Ava_9-21hrs(Interval Count)',index=False)     
    df5.to_excel(writer, sheet_name='LTE_TXN_Link_Capacity',index=False)
    df7.to_excel(writer, sheet_name='S1_HO_Attempts',index=False)
    df8.to_excel(writer, sheet_name='Flash_CSFB',index=False)
    df9.to_excel(writer, sheet_name='Base_Station_Transport_Data',index=False)
    df10.to_excel(writer, sheet_name='Maximum_User_Number',index=False)
    df11.to_excel(writer, sheet_name='LTE_eNodeB_SON=OFF',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
