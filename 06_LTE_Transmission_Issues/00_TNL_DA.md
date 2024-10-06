#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

## Delete Files From the Folder, If Folder Not Found Created the Folder


```python
import os

# Folder path
folder_path = r"D:/Advance_Data_Sets/BSS_EI_Output/TNL_Output"

# Check if the folder exists
if os.path.exists(folder_path):
    # Iterate over the files in the folder and delete them
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted file: {file_path}")
        except Exception as e:
            print(f"Error deleting {file_path}: {e}")
else:
    # If the folder doesn't exist, create it
    print(f"Folder {folder_path} does not exist. Creating...")
    os.makedirs(folder_path)
    print(f"Folder {folder_path} created.")
```

# LTE TXN & Availability Issues (DA)

## Import Required Libraries


```python
# Required Libraries
import os
import numpy as np
import pandas as pd
from glob import glob

import warnings
warnings.simplefilter("ignore")
```

## Set Input File 


```python
# set the input file path
os.chdir('D:/Advance_Data_Sets/TXN_DataSets/LTE_TNL_BTS_DA')
```

## Import & Concat Data Set


```python
# List of Zip Files in a given Path
df0 = sorted(glob('*.zip'))
# import & concat all the Cell DA KPIs
df=pd.concat((pd.read_csv(file,skiprows=range(6),\
    skipfooter=1,engine='python',
    parse_dates=["Date"],na_values=['NIL','/0'],\
                usecols=['Date','eNodeB Name','L.E-RAB.FailEst.TNL',
                        'L.E-RAB.AbnormRel.TNL',
                        'Radio Network Unavailability Rate_Cell',
                        'S1 Handover Success Rate_DEN']) for file in df0)).\
                        sort_values('Date').set_index(['Date']).\
                        last('7D').reset_index()
```

# L.E-RAB.FailEst.TNL (DA)


```python
# Re-Shape Data Set
df1 = df.pivot_table(index='eNodeB Name',columns='Date',values='L.E-RAB.FailEst.TNL').reset_index()

# Conditional Count For P0
df1["Number_of_Days_TNL>=500"] = df1.where(df1.iloc[:,1:].\
                                  ge(500)).count(axis=1)

# Conditional Count For P1
df1["Number_of_Days_200<TNL<500"] = df1.where((df1.iloc[:,1:-1].ge(200)) &\
                                              (df1.iloc[:,1:-1].lt(500))).count(axis=1)

# Assagin Priority
df1['LTE_TNL_Aging_Priority'] = np.where(
            (df1['Number_of_Days_TNL>=500'].ge(4)),
            'P0', 
            np.where(
                    (df1['Number_of_Days_200<TNL<500'].ge(4)), 
                    'P1',
                     '-'))
```

## L.E-RAB.AbnormRel.TNL (DA)


```python
# Re-Shape Data Set
df2 = df.pivot_table(index='eNodeB Name',columns='Date',values='L.E-RAB.AbnormRel.TNL').reset_index()

# Conditional Count For P0
df2["Number_of_Days_Abnormrel_TNL>=500"] = df2.where(df2.iloc[:,1:].\
                                  ge(500)).count(axis=1)

# Conditional Count For P1
df2["Number_of_Days_200<Abnormrel_TNL<500"] = df2.where((df2.iloc[:,1:-1].ge(200)) &\
                                              (df2.iloc[:,1:-1].lt(500))).count(axis=1)

# Assagin Priority
df2['LTE_Abnormrel_TNL_Aging_Priority'] = np.where(
            (df2['Number_of_Days_Abnormrel_TNL>=500'].ge(4)),
            'P0', 
            np.where(
                    (df2['Number_of_Days_200<Abnormrel_TNL<500'].ge(4)), 
                    'P1',
                     '-'))
```

## Radio Network Unavailability Rate_Cell


```python
# Re-Shape Data Set
df3 = df.pivot_table(index='eNodeB Name',columns='Date',values='Radio Network Unavailability Rate_Cell').reset_index()
# Conditional Count
df3["Number_of_Days_Unavailability>0"] = df3.where(df3.iloc[:,1:].gt(0)).count(axis=1)
```

## S1 HO Attempts 


```python
# Re-Shape Data Set
df4 = df.pivot_table(index='eNodeB Name',columns='Date',values='S1 Handover Success Rate_DEN').reset_index()
# Conditional Count
df4["S1_HO_Attempts>=20"] = df4.where(df4.iloc[:,1:].ge(20)).count(axis=1)
```

## Export Output


```python
# set the input file path
os.chdir('D:/Advance_Data_Sets/BSS_EI_Output/TNL_Output')
```


```python
# Write excel file with default behaviour.
with pd.ExcelWriter("00_LTE_TXN_Availability_Issues.xlsx",date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    df1.to_excel(writer, sheet_name='LTE_TNL_Fail',index=False)
    df2.to_excel(writer, sheet_name='LTE_AbnormRel_TNL_Fail',index=False)
    df3.to_excel(writer, sheet_name='LTE_Availability_Issues',index=False) 
    df4.to_excel(writer, sheet_name='S1_HO_Attempts',index=False) 
```


```python
# re-set all the variable from the RAM
%reset -f
```
