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
import os
#import glob
import zipfile
import pandas as pd
from glob import glob
from datetime import datetime, timedelta
import warnings
warnings.simplefilter("ignore")
```

## Set Input File Path


```python
# Set Export File Path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/2G_Outage_Intervals_BTS_Level'
os.chdir(folder_path)
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

## Import and Concat Outage Interval Files


```python
# concat all the BTS Hourly Files
df=pd.concat((pd.read_csv(file, engine='python',\
            parse_dates=["Date"],na_values=['NIL','/0']) for file in filtered_file_list)).\
            sort_values('Date').reset_index()
```

## re-shape data set (00:00 hrs - 23:00 hrs)


```python
# re-shape data set (00:00 hrs - 23:00 hrs)
df0 = pd.pivot_table(df,index=['GBSC','Site Name'],\
                     columns='Date',\
                     values='Total Count of (TCH Ava Rate)<100 between 0:00-23:00',\
                     margins=True,\
                     margins_name='Total Intervals (TCH Ava)<100 between 0:00-23:00',\
                     aggfunc=sum).\
                     reset_index().\
                     iloc[:-1, :]
```

## re-shape data set (09:00 hrs - 21:00 hrs)


```python
# re-shape data set (09:00 hrs - 21:00 hrs)
df1= pd.pivot_table(df,index=['GBSC','Site Name'],\
                    columns='Date',values='Total Count of (TCH Ava Rate)<100 between 9:00-21:00',\
                    margins=True,\
                    margins_name='Total Intervals (TCH Ava)<100 between 9:00-21:00',
                    aggfunc=sum).\
                    reset_index().\
                    iloc[:-1, :]
```

## Export Output


```python
folder_path= 'D:/Advance_Data_Sets/BSS_EI_Output/CDIG_Processing'
os.chdir(folder_path)

with pd.ExcelWriter('02_CDIG_Analysis.xlsx',date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    df0.to_excel(writer,sheet_name="GSM_Ava_24hrs_Interval_Count",engine='openpyxl',index=False)
    df1.to_excel(writer,sheet_name="GSM_Ava_9-21hrs_Interval_Count",engine='openpyxl',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
