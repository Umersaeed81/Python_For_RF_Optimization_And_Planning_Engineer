#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# Special Characters Audit

## Import Libraries


```python
import os
import pandas as pd
import warnings
warnings.simplefilter("ignore")
```

## Set Input File Path


```python
# set the CME Export Input File Path
path_cell_hourly = 'D:/Advance_Data_Sets/CME_Export'
os.chdir(path_cell_hourly)
```

## User Define Funcation


```python
# Define a function to load Excel sheets and check for special characters
def load_and_check(sheet_name, cols):
    df = pd.read_excel('CME_North.xlsx', sheet_name=sheet_name, header=1, usecols=cols)
    for col in cols:
        df[f'{col}_Check'] = df[col].str.contains(r'^\s|[^\a-zA-Z0-9_()/[&#.\- ]|\s$| {2,}')
    return df
```


```python
# List of sheet names and columns to read
sheets_and_columns = [
    ('GCELL', ['BSC Name','BTS Name', 'Cell Name']),
    ('GCELLFREQ_FREQ', ['BSC Name','Cell Name']),
    ('GTRX', ['BSC Name','Cell Name', 'TRX Name']),
    ('GTRXHOP', ['BSC Name','Cell Name', 'TRX Name']),
    ('GTRXCHANHOP', ['BSC Name','Cell Name', 'TRX Name']),
    ('GCELLMAGRP', ['BSC Name','Cell Name']),
    ('TRXBIND2PHYBRD', ['BSC Name','Cell Name', 'TRX Name', 'BTS Name'])
]
```


```python
# Load data and check for special characters
dfs = {}
for sheet_name, cols in sheets_and_columns:
    dfs[sheet_name] = load_and_check(sheet_name, cols)
```


```python
# Extract individual DataFrames for further processing if needed
df0 = dfs['GCELL']
df1 = dfs['GCELLFREQ_FREQ']
df2 = dfs['GTRX']
df3 = dfs['GTRXHOP']
df4 = dfs['GTRXCHANHOP']
df5 = dfs['GCELLMAGRP'].dropna()
df6 = dfs['TRXBIND2PHYBRD']
```

## Export Output


```python
# Write excel file with default behaviour.
with pd.ExcelWriter("CME_Check_List.xlsx") as writer:
    df0.to_excel(writer, sheet_name='GCELL',index=False)
    df1.to_excel(writer, sheet_name='GCELLFREQ_FREQ',index=False)
    df2.to_excel(writer, sheet_name='GTRX',index=False)    
    df3.to_excel(writer, sheet_name='GTRXHOP',index=False)
    df4.to_excel(writer, sheet_name='GTRXCHANHOP',index=False)     
    df5.to_excel(writer, sheet_name='GCELLMAGRP',index=False)
    df6.to_excel(writer, sheet_name='TRXBIND2PHYBRD',index=False)
```
