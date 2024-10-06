#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# Evaluate the Transmission Link Capacity

## Data Processing Steps

1. Import and Concatenate All the csv Files (Hourly KPIs).
2. Calculate $max(VS.FEGE.RxMaxSpeed(bit/s))*0.97$
3. Then Calculate the **difference** between **Step-2 value** & $VS.FEGE.RxMaxSpeed(bit/s)$
4. Calculate the **Number of intervals for each date** where difference is less than equal to  $0$ in pivot table format.
5. conditional days count (if intervals >=7).
6. Filter the days >=4 are consider the congested eNodeB.

## Import Required Libraries


```python
# Required Libraries
import os
import numpy as np
import pandas as pd
from glob import glob
from datetime import datetime, timedelta
```

## Input File Path


```python
# set the input file path
os.chdir('D:/Advance_Data_Sets/TXN_DataSets/LTE_Transmission_Link_Capacity')
```


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

## Import and Concat Input File path


```python
# import & concat all the Cell Hourly KPIs
df0=pd.concat((pd.read_csv(file,skiprows=range(6),\
    skipfooter=1,engine='python',
    parse_dates=["Date"],na_values=['NIL','/0']) for file in filtered_file_list))\
    .sort_values('Date').dropna().reset_index()
```

## Required Calculation


```python
"""
Calculate max value of 'VS.FEGE.RxMaxSpeed(bit/s)' 
for each BTS and Date and multiply with 0.97
"""

df0['max'] = (df0.groupby(['Date','eNodeB Name'])\
             [['VS.FEGE.RxMaxSpeed(bit/s)']].transform('max'))*0.97


# Calculate the difference
df0['difference'] = df0['max']-df0['VS.FEGE.RxMaxSpeed(bit/s)']

# drop un-required columns
df0 = df0.drop(['max','Integrity'], axis=1)
```

## Conditional pivot table and formatting


```python
# Conditional pivot table
df1=df0.loc[df0['difference'].le(0)]\
            .pivot_table(index=['eNodeB Name'],\
            columns='Date',values=['difference'],\
            aggfunc='count',
            margins=True,
            margins_name='Total_Intervals_Difference<=0')\
            .fillna(0)\
            .reset_index()
# drop last row of the pivot_table
df1=df1.iloc[:-1,:]

# drop level zero header
df1.columns = df1.columns.droplevel(0)

# re-name the header
df1= df1.rename(columns = {'':'eNodeB Name'})
```

## Conditional Counts


```python
# conditional days count 
df1["#of_Days_Difference<=0"] = df1.\
        where(df1.iloc[:,1:-1] >= 7)\
        .count(axis=1)

# filter requied rows
df2=df1[df1["#of_Days_Difference<=0"] >=4].\
            copy().reset_index(drop=True)
```

## Export Output


```python
# set the input file path
os.chdir('D:/Advance_Data_Sets/BSS_EI_Output/TNL_Output')
```


```python
# Write excel file with default behaviour.
with pd.ExcelWriter("02_LTE_TXN_Availability_Issues.xlsx",date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    df2.to_excel(writer, sheet_name='LTE_TXN_Link_Capacity',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```

