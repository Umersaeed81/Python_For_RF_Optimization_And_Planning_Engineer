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
from datetime import datetime, timedelta
import warnings
warnings.simplefilter("ignore")
```

## Set Input File Path (TNL File Without Outage)


```python
# set BTS Hourly file Path
os.chdir('D:/Advance_Data_Sets/BSS_EI_Output/TNL_Ava_Interval/TNL')
```

## Filter Required Files (TNL File Without Outage)


```python
# List all files in the path
file_list = glob('*.csv')

# Calculate the date for three days ago
required_date = datetime.now() - timedelta(days=7)
required_date_str = required_date.strftime("%d%m%Y")

# Filter files that match the required date or the two days before
filtered_file_list = [file for file in file_list if any((required_date + timedelta(days=i)).strftime("%d%m%Y") in file for i in range(7))]

if filtered_file_list:
    print("Filtered File List:", filtered_file_list)
else:
    print(f"No file found with the required dates in the last three days.")
```

## Import & Concat Data Set (TNL Files)


```python
# import & concat all the Hourly KPIs
df=pd.concat((pd.read_csv(file,\
                          engine='python',\
                          parse_dates=["Date"],\
                          na_values=['NIL','/0']) for file in filtered_file_list)).\
                          sort_values('Date').reset_index(drop=True)
```

## Re-shapre Date Set (pivot_table)


```python
df1 = df.pivot_table(index='eNodeB Name',columns='Date',values='L.E-RAB.FailEst.TNL').reset_index().fillna(0)
```

## Assagin Priority


```python
# Conditional Count For P0
df1["Number_of_Days_TNL>=500"] = df1.where(df1.iloc[:,1:].\
                                  ge(500)).count(axis=1)

# Conditional Count For P1
df1["Number_of_Days_200<TNL<500"] = df1.where((df1.iloc[:,1:-1].ge(200)) &\
                                              (df1.iloc[:,1:-1].lt(500))).count(axis=1)

```

## Assagin Priority


```python
df1['LTE_TNL_Aging_Priority'] = np.where(
            (df1['Number_of_Days_TNL>=500'].ge(4)),
            'P0', 
            np.where(
                    (df1['Number_of_Days_200<TNL<500'].ge(4)), 
                    'P1',
                     '-'))
```

## Set Input File Path (Outage Interval)


```python
# set BTS Hourly file Path
os.chdir('D:/Advance_Data_Sets/BSS_EI_Output/TNL_Ava_Interval/Availability')
```


```python
# List all files in the path
file_list = glob('*.csv')

# Calculate the date for three days ago
required_date = datetime.now() - timedelta(days=7)
required_date_str = required_date.strftime("%d%m%Y")

# Filter files that match the required date or the two days before
filtered_file_list = [file for file in file_list if any((required_date + timedelta(days=i)).strftime("%d%m%Y") in file for i in range(7))]

if filtered_file_list:
    print("Filtered File List:", filtered_file_list)
else:
    print(f"No file found with the required dates in the last three days.")
```

## Import & Concat Data Set (Outage Files)


```python
# import & concat all the Hourly KPIs
df2=pd.concat((pd.read_csv(file,\
                          engine='python',\
                          parse_dates=["Date"],\
                          na_values=['NIL','/0']) for file in filtered_file_list)).\
                          sort_values('Date').reset_index(drop=True)
```

# re-shape data set (00:00 hrs - 23:00 hrs)


```python
df3= pd.pivot_table(df2,index=['eNodeB Name'],\
                    columns='Date',values='Total Count of (Down Time)>0 between 0:00-23:00',\
                    margins=True,\
                    margins_name='Total Intervals (Down Time)>0 between 0:00-23:00',
                    aggfunc=sum).\
                    reset_index().\
                    iloc[:-1, :]
```

## re-shape data set (09:00 hrs - 21:00 hrs)


```python
df4= pd.pivot_table(df2,index=['eNodeB Name'],\
                    columns='Date',values='Total Count of (Down Time)>0 between 9:00-21:00',\
                    margins=True,\
                    margins_name='Total Intervals (Down Time)>0 between 9:00-21:00',
                    aggfunc=sum).\
                    reset_index().\
                    iloc[:-1, :]
```

## Output


```python
# set the input file path
os.chdir('D:/Advance_Data_Sets/BSS_EI_Output/TNL_Output')

# Write excel file with default behaviour.
with pd.ExcelWriter("01_LTE_TXN_Availability_Issues.xlsx",date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    df1.to_excel(writer, sheet_name='LTE_Pure_TNL',index=False)
    df3.to_excel(writer, sheet_name='LTE_Ava_24hrs(Interval Count)',index=False)
    df4.to_excel(writer, sheet_name='LTE_Ava_9-21hrs(Interval Count)',index=False)  
```
