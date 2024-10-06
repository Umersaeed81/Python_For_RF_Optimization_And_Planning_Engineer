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

## UMTS Cell Hourly KPIs File Path


```python
# Set Path for GSM Cell DA KPIs
path_cell_hourly = 'D:/Advance_Data_Sets/BSS_EI_Output/3G_EI_Ava_Intervals'
os.chdir(path_cell_hourly)
```

## Filter Required Files


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

    Filtered File List: ['290_UMTS_RTWP_DownTime_Intervals_27092024.csv', '291_UMTS_RTWP_DownTime_Intervals_28092024.csv', '292_UMTS_RTWP_DownTime_Intervals_29092024.csv', '293_UMTS_RTWP_DownTime_Intervals_30092024.csv', '294_UMTS_RTWP_DownTime_Intervals_01102024.csv', '295_UMTS_RTWP_DownTime_Intervals_02102024.csv', '296_UMTS_RTWP_DownTime_Intervals_03102024.csv']
    

## Import and Concat Input Files


```python
df=pd.concat((pd.read_csv(file,\
    skipfooter=1,engine='python',
            parse_dates=["Date"],na_values=['NIL','/0'],\
             usecols=['Date', 'RNC', 'Cell ID',
               'Total Count of (Down Time)>0 between 0:00-23:00',
               'Total Count of (Down Time)>0 between 9:00-21:00'],
            dtype = {"Cell ID" : str}) for file in filtered_file_list))\
            .sort_values('Date').set_index(['Date']).\
             last('7D').reset_index()
```


```python
df.head(2)
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Date</th>
      <th>RNC</th>
      <th>Cell ID</th>
      <th>Total Count of (Down Time)&gt;0 between 0:00-23:00</th>
      <th>Total Count of (Down Time)&gt;0 between 9:00-21:00</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>2024-09-27</td>
      <td>FSDRNC02</td>
      <td>10043</td>
      <td>8</td>
      <td>5</td>
    </tr>
    <tr>
      <th>1</th>
      <td>2024-09-27</td>
      <td>PSHRNC01</td>
      <td>22668</td>
      <td>4</td>
      <td>0</td>
    </tr>
  </tbody>
</table>
</div>



## Re-Shape (Pivot_table)


```python
# re-shape data set (00:00 hrs - 23:00 hrs)
df0 = pd.pivot_table(df,index=['RNC','Cell ID'],\
                     columns='Date',\
                     values='Total Count of (Down Time)>0 between 0:00-23:00',\
                     margins=True,\
                     margins_name='Total Intervals (Down Time)>0 between 0:00-23:00',\
                     aggfunc=sum).\
                     reset_index().\
                     iloc[:-1, :]
```


```python
# re-shape data set (09:00 hrs - 21:00 hrs)
df1= pd.pivot_table(df,index=['RNC','Cell ID'],\
                    columns='Date',values='Total Count of (Down Time)>0 between 9:00-21:00',\
                    margins=True,\
                    margins_name='Total Intervals (Down Time)>0 between 9:00-21:00',
                    aggfunc=sum).\
                    reset_index().\
                    iloc[:-1, :]
```

## Export Data Set


```python
# set path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files'
os.chdir(folder_path)

# Write excel file with default behaviour.
with pd.ExcelWriter("03_UMTS_BSS_Issus_Intervals.xlsx",date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    # UMTS BSS Issues
    df0.to_excel(writer, sheet_name='UMTS_Ava_24hrs Interval Count',index=False)
    df1.to_excel(writer, sheet_name='UMTS_Ava_9-21hrs Interval Count',index=False)
```


```python
#re-set all the variable from the RAM
%reset -f
```


```python

```


```python

```
