#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# External Interference KPIs (3G Intervals in Pivot table Format)

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

## Filter Required Files


```python
# List all files in the path
file_list = glob('D:/Advance_Data_Sets/BSS_EI_Output/3G_EI_Ava_Intervals/*.csv')

# Calculate the date for three days ago
required_date = datetime.now() - timedelta(days=10)
required_date_str = required_date.strftime("%d%m%Y")

# Filter files that match the required date or the two days before
filtered_file_list = [file for file in file_list if any((required_date + timedelta(days=i)).strftime("%d%m%Y") in file for i in range(10))]

if filtered_file_list:
    print("Filtered File List:", filtered_file_list)
else:
    print(f"No file found with the required dates in the last three days.")
```

## Import Required Files


```python
df=pd.concat((pd.read_csv(file,engine='python',na_values=['NIL','/0'],\
    parse_dates=["Date"]) for file in filtered_file_list))
```

## Re-shape Data Set (Pivot table)


```python
# per day intervals (pivot_table format)
df0= df.pivot_table(index=['RNC','Cell ID'],\
                    columns="Date",\
                     values='Total_Interval_RTWP>=-95',\
                     margins=True,\
                     aggfunc=sum)\
                    .reset_index()\
                    .iloc[:-1, :]\
                    .rename({'All': 'Total_Interval_Impacted_RTWP>=-95'}, axis=1)    
```


```python
# per day intervals (pivot_table format)
df1= df.pivot_table(index=['RNC','Cell ID'],\
                    columns="Date",\
                     values='Total_Interval_RTWP>=-98',\
                     margins=True,\
                     aggfunc=sum)\
                    .reset_index()\
                    .iloc[:-1, :]\
                    .rename({'All': 'Total_Interval_Impacted_RTWP>=-98'}, axis=1)   
```

## Export Output


```python
# Set Output File Path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files_EI'
os.chdir(folder_path)

with pd.ExcelWriter('04_External_Interference_UMTS_Intervals.xlsx',date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    # UMTS EI Count Pivot table format (RTWP>=95)
    df0.to_excel(writer,sheet_name="UMTS_RTWP_Intervals_-95",index=False)
    # UMTS EI Count Pivot table format (RTWP>=98)
    df1.to_excel(writer,sheet_name="UMTS_RTWP_Intervals_-98",index=False)
```


```python
# Set Output File Path
folder_path = 'D:/Advance_Data_Sets/Output_Folder'
os.chdir(folder_path)
df.to_excel('RTWP_Intervals.xlsx',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
