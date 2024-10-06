#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# BSS & EI Issues (UMTS Hourly)

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

## UMT Cell Hourly KPIs File Path


```python
# Set Path for UMTS Cell DA KPIs
path_cell_hourly = 'D:/Advance_Data_Sets/KPIs_Analysis/Hourly_KPIs_Cell_Level/UMTS'
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
                       usecols=['Date','Time','RNC','Cell ID',\
                      'Cell Unavailability duration (Down Time)','VS.MeanRTWP(dBm)'],
                      parse_dates=["Date","Time"], na_values=['NIL','/0'],
                      dtype={"Cell ID" : int,'RNC':str,\
                        'Cell Unavailability duration (Down Time)':float,\
                         'VS.MeanRTWP(dBm)':float})\
                      for file in filtered_file_list)
```




## Calculate Degraded Intervals


```python
# Calculate intervals
df1 = df0.groupby(['Date','RNC', 'Cell ID']).apply(lambda x: pd.Series({
        'Total Count of (Down Time)>0 between 0:00-23:00': (x['Cell Unavailability duration (Down Time)'].gt(0)).sum(),
        'Total Count of (Down Time)>0 between 9:00-21:00': (x.set_index("Time").between_time('9:00', '21:00')['Cell Unavailability duration (Down Time)'].gt(0)).sum(),   
        'Total_Interval_RTWP>=-95': (x['VS.MeanRTWP(dBm)'].ge(-95)).sum(),
        'Total_Interval_RTWP>=-98': (x['VS.MeanRTWP(dBm)'].ge(-98)).sum()
    })).reset_index()
```




## Idendtify the Max RTWP for UMTS


```python
df2 = df0[['Date','Time','RNC','Cell ID','VS.MeanRTWP(dBm)']].dropna().reset_index(drop=True)
```





```python
df3 = df2.loc[df2.groupby(['Date','RNC','Cell ID'])\
            ['VS.MeanRTWP(dBm)'].idxmax()].reset_index(drop=True)
```

## Export Max RTWP KPIs (UMTS)


```python
# Set Export File Path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/3G_EI_Max'
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
    output_file = f'{new_max_num:02d}_UMTS_RTWP_Max_{unique_date_str}.csv'
    
    # Filter the DataFrame for the current unique date
    df_filtered = df3[df3['Date'] == unique_date]
    
    # Extract the time portion from the 'Time' column and convert to 24-hour format
    df_filtered['Time'] = df_filtered['Time'].dt.strftime('%H:%M')
    
    # Export to CSV
    df_filtered.to_csv(output_file, index=False)

    # Increment the maximum number for the next iteration
    new_max_num += 1
```

## Export Down Time and RTWP Degraded Intervals


```python
# Set Export File Path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/3G_EI_Ava_Intervals'
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
    output_file = f'{new_max_num:02d}_UMTS_RTWP_DownTime_Intervals_{unique_date_str}.csv'
    
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



