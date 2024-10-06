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
import shutil
import zipfile
import numpy as np
import pandas as pd
from glob import glob
from datetime import datetime, timedelta
import warnings
warnings.simplefilter("ignore")
```

## Set Input File Path


```python
# Set Cell Hourly KPIs Path 
os.chdir('D:/Advance_Data_Sets/KPIs_Analysis/Hourly_KPIs_Cell_Level/LTE')
```

## Filter Required Files


```python
# List all files in the path
file_list = glob('*.zip')

# Calculate the date for three days ago
required_date = datetime.now() - timedelta(days=1)
required_date_str = required_date.strftime("%d%m%Y")

# Filter files that match the required date or the two days before
filtered_file_list = [file for file in file_list if any((required_date + timedelta(days=i)).strftime("%d%m%Y") in file for i in range(1))]

if filtered_file_list:
    print("Filtered File List:", filtered_file_list)
else:
    print(f"No file found with the required dates in the last three days.")
```

## Import and concat Input Files


```python
df0=pd.concat((pd.read_csv(file,skiprows=range(6),
        skipfooter=1,engine='python',na_values=['NIL','/0'],
        parse_dates=["Date"],usecols=['Date','Time','Cell Name',
            'L.ChMeas.PRB.UL.Avail','L.ChMeas.PRB.DL.Avail']) for file in filtered_file_list)).\
             sort_values('Date').dropna().sort_values(['Date', 'Cell Name']).reset_index(drop=True)
```

## Data Pre-Processing


```python
# Step 2: Clean 'Cell Name' column thoroughly
df0['Cell Name'] = df0['Cell Name'].str.strip()  # Remove leading/trailing spaces
df0['Cell Name'] = df0['Cell Name'].str.replace(r'\s+', ' ', regex=True)  # Replace multiple spaces with single space
df0['Cell Name'] = df0['Cell Name'].str.replace('\xa0', ' ', regex=False)  # Replace non-breaking spaces if any
```

## Calculate Per Day Max DL PRBs


```python
df1 = df0.loc[df0.groupby(['Date','Cell Name'])['L.ChMeas.PRB.DL.Avail'].idxmax()].reset_index(drop=True).drop('L.ChMeas.PRB.UL.Avail', axis=1)
```

## Calculate Per Day Max UL PRBs


```python
df2 = df0.loc[df0.groupby(['Date','Cell Name'])['L.ChMeas.PRB.UL.Avail'].idxmax()].reset_index(drop=True).drop('L.ChMeas.PRB.DL.Avail', axis=1)
```


```python
df3 = pd.merge(df1[['Date','Cell Name','L.ChMeas.PRB.DL.Avail']],df2[['Date','Cell Name','L.ChMeas.PRB.UL.Avail']],on=['Date','Cell Name'])
```
