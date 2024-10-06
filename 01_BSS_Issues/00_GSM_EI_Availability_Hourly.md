#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# BSS & EI Issues (GSM Hourly)

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

## GSM Cell Hourly KPIs File Path


```python
# Set Path for GSM Cell DA KPIs
path_cell_hourly = 'D:/Advance_Data_Sets/KPIs_Analysis/Hourly_KPIs_Cell_Level/GSM'
os.chdir(path_cell_hourly)
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

    Filtered File List: ['GSM_Cell_Houlry_03102024.zip']
    

## Import and Concat Cell Hourly KPIs


```python
# Define a function to read csv files from zip archives
def read_csv_from_zip(file, **kwargs):
    with zipfile.ZipFile(file) as z:
        # Get list of csv files in the zip archive
        csv_files = [name for name in z.namelist() if name.endswith('.csv')]
        # Concatenate all csv files into a single dataframe
        dfs = [pd.read_csv(z.open(name), **kwargs) for name in csv_files]
        return pd.concat(dfs, ignore_index=True)
```


```python
# Concatenate all csv files from all zip files
df0 = pd.concat(
    read_csv_from_zip(file, skiprows=range(6), skipfooter=1, engine='python',
                       usecols=['Date','Time','GBSC','Site Name','Cell CI',\
                      'TCH Availability Rate(%)','Interference Band Proportion (4~5)(%)'],
                      parse_dates=["Date","Time"], na_values=['NIL','/0'],
                      dtype={"Cell CI" : int,'Site Name':str,'TCH Availability Rate(%)':float})\
                      for file in filtered_file_list)
```


```python
df0.head(2)
```

<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Date</th>
      <th>Time</th>
      <th>GBSC</th>
      <th>Cell CI</th>
      <th>Site Name</th>
      <th>TCH Availability Rate(%)</th>
      <th>Interference Band Proportion (4~5)(%)</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>2024-10-03</td>
      <td>2024-10-04</td>
      <td>HLHRBSC01</td>
      <td>17625</td>
      <td>7625_Sunder Road Raiwind City Kasur (3G-CI-4984)</td>
      <td>100.0</td>
      <td>0.0289</td>
    </tr>
    <tr>
      <th>1</th>
      <td>2024-10-03</td>
      <td>2024-10-04</td>
      <td>HLHRBSC01</td>
      <td>37625</td>
      <td>7625_Sunder Road Raiwind City Kasur (3G-CI-4984)</td>
      <td>100.0</td>
      <td>0.0000</td>
    </tr>
  </tbody>
</table>
</div>



## Calculate Degraded Intervals


```python
# Calculate Low Availability Intervals
df1 = df0.groupby(['Date','GBSC','Site Name','Cell CI']).apply(lambda x: pd.Series({
        'Total Count of Intervals (TCH Availability Rate)<100 between 0:00-23:00': (x['TCH Availability Rate(%)'].lt(100)).sum(),
        'Total Count of Intervals (TCH Availability Rate)<100 between 9:00-21:00': (x.set_index("Time").between_time('9:00', '21:00')['TCH Availability Rate(%)'].lt(100)).sum(),
        'Total Count of Intervals Interference Band Proportion (4~5)(%)>10': (x['Interference Band Proportion (4~5)(%)'].ge(10)).sum()
        })).reset_index()
```

## Idendtify the max Interference for GSM


```python
df2 = df0[['Date','Time','GBSC','Cell CI','Site Name','Interference Band Proportion (4~5)(%)']].dropna().reset_index(drop=True)
```


```python
df3 = df2.loc[df2.groupby(['Date','GBSC','Cell CI','Site Name'])\
            ['Interference Band Proportion (4~5)(%)']\
              .idxmax()].reset_index(drop=True)
```

## Export Max EI KPIs (GSM)


```python
# Set Export File Path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/2G_EI_Max'
os.chdir(folder_path)
files = os.listdir()

# Check if files list is empty
if not files:
    new_max_num = 0
else:
    # Extract the maximum number from existing files
    max_num = max(int(file.split('_')[0]) for file in files)
    # Increment the maximum number
    new_max_num = max_num + 1

# Iterate over unique dates and export to separate CSV files
for unique_date in df3['Date'].unique():
    unique_date_str = unique_date.strftime('%d%m%Y')
    
    # Generate the output file name
    output_file = f'{new_max_num:02d}_GSM_EI_Max_{unique_date_str}.csv'
    
    # Filter the DataFrame for the current unique date
    df_filtered = df3[df3['Date'] == unique_date]
    
    # Extract the time portion from the 'Time' column and convert to 24-hour format
    df_filtered['Time'] = df_filtered['Time'].dt.strftime('%H:%M')
    
    # Export to CSV
    df_filtered.to_csv(output_file, index=False)

    # Increment the maximum number for the next iteration
    new_max_num += 1
```

## Export TCH Availability and EI Degraded Intervals


```python
# Set Export File Path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/2G_EI_Ava_Intervals'
os.chdir(folder_path)
files = os.listdir()

# Check if files list is empty
if not files:
    new_max_num = 0
else:
    # Extract the maximum number from existing files
    max_num = max(int(file.split('_')[0]) for file in files)
    # Increment the maximum number
    new_max_num = max_num + 1

# Iterate over unique dates and export to separate CSV files
for unique_date in df1['Date'].unique():
    unique_date_str = unique_date.strftime('%d%m%Y')
    
    # Generate the output file name
    output_file = f'{new_max_num:02d}_GSM_EI_Availability_Intervals_{unique_date_str}.csv'
    
    # Filter the DataFrame for the current unique date
    df_filtered = df1[df1['Date'] == unique_date]
        
    # Export to CSV
    df_filtered.to_csv(output_file, index=False)

    # Increment the maximum number for the next iteration
    new_max_num += 1
```


```python
#re-set all the variable from the RAM
%reset -f
```
