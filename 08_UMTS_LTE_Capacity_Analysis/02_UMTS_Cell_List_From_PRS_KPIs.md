#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# UMTS Tech Cell List From KPIs

## Import Required Libraries


```python
# Import Libraries
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

## Filter UMTS KPIs Files


```python
# List all files in the path
file_list = glob('D:/Advance_Data_Sets/GUL/3G/*.zip')

# Calculate the date for 14 days ago
required_date = datetime.now() - timedelta(days=14)
required_date_str = required_date.strftime("%d%m%Y")

# Filter files that match the required date or the two days before
filtered_file_list = [file for file in file_list if any((required_date + timedelta(days=i)).strftime("%d%m%Y") in file for i in range(14))]

if filtered_file_list:
    print("Filtered File List:", filtered_file_list)
else:
    print(f"No file found with the required dates in the last 14 days.")
```

## Import & Concat UMTS Cell Busy KPIs


```python
# import & concat all the Cell Busy KPIs
df=pd.concat((pd.read_csv(file,skiprows=range(6),\
    skipfooter=1,engine='python',
    parse_dates=["Date"],na_values=['NIL','/0'],\
                usecols=['Date','NODEBNAME','CELLNAME']) for file in filtered_file_list)).\
                sort_values('Date').set_index(['Date']).last('14D').reset_index()
```

## Replace multiple spaces with single space in 'eNodeB Name' & 'CELLNAME' column


```python
# Replace multiple spaces with single space in 'Cell Name' column
df['CELLNAME'] = df['CELLNAME'].str.replace(r'\s+', ' ', regex=True)
# Replace multiple spaces with single space in 'NodeB Name' column
df['NODEBNAME'] = df['NODEBNAME'].str.replace(r'\s+', ' ', regex=True)
```

## Replace Cell Name (atleast the audit once in a mounth)


```python
df['CELLNAME'] = df['CELLNAME'].replace({
    '3G-S-2654_Kashki Hamza Zai (USF) (S-7424)': '3G-S-2654-3_Kashki Hamza Zai (USF) (S-7424)',
    '3G-N-4243-K_Kot Azmath (Gold) Bannu (N-4757)': '3G-N-4243-4_Kot Azmath (Gold) Bannu (N-4757)'
})
```

## Filter Required Columns and remove duplicates


```python
df1 = df [['CELLNAME','NODEBNAME']].drop_duplicates()
```

## Generate UMTS Site/Cell IDs


```python
# Generate UMTS Site ID
df1['UMTS_Site_ID'] = df1['NODEBNAME'].str.split('_').str[0]
# Generate UMTS Cell ID
df1['UMTS_Cell_ID'] = df1['CELLNAME'].str.split('_').str[0]
```

## Add Comment Column


```python
df1['Comments'] = 'Yes'
```

## Export Output


```python
# Set Output Folder Path
folder_path = 'D:/Advance_Data_Sets/GUL/GUL_Output'
os.chdir(folder_path)
# export
df1.to_csv('01_UMTS_KPIs_Cell_List.csv',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
