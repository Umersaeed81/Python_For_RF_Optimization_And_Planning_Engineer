#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# Flash_CSFB (%)

## Import Required Libraries


```python
# Required Libraries
import os
import numpy as np
import pandas as pd
from glob import glob
from datetime import datetime, timedelta
```

## Set Input File 


```python
# set the input file path
os.chdir('D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/LTE')
```

## Filter the Required Files


```python
# List all files in the path
file_list = glob('*.zip')

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

## Import & Concat Data Set


```python
# List of Zip Files in a given Path
#df0 = sorted(glob('*.zip'))
# import & concat all the Cell DA KPIs
df=pd.concat((pd.read_csv(file,skiprows=range(6),\
    skipfooter=1,engine='python',
    parse_dates=["Date"],na_values=['NIL','/0'],\
                usecols=['Date','eNodeB Name','Cell Name',
                'LocalCell Id','eNodeB Function Name',
                'Flash CSFB (%)']) for file in filtered_file_list)).reset_index(drop=True)
```

## Flash CSFB (%)


```python
# Re-Shape Data Set
df1 = df.pivot_table(index=['eNodeB Name','Cell Name',
                'LocalCell Id','eNodeB Function Name'],\
                     columns='Date',values='Flash CSFB (%)').reset_index()
```


```python
df1["Flash CSFB=0"] = df1.where(df1.iloc[:,4:].eq(0)).count(axis=1)
```

## Export Output


```python
# set the input file path
os.chdir('D:/Advance_Data_Sets/BSS_EI_Output/TNL_Output')
```


```python
# Write excel file with default behaviour.
with pd.ExcelWriter("03_LTE_TXN_Availability_Issues.xlsx",date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    df1.to_excel(writer, sheet_name='Flash_CSFB',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
