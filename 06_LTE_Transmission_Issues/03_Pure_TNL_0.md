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
# Required Libraries
import os
import numpy as np
import pandas as pd
from glob import glob
from datetime import datetime, timedelta
import warnings
warnings.simplefilter("ignore")
```

## Set Input File Path


```python
# set BTS Hourly file Path
os.chdir('D:/Advance_Data_Sets/TXN_DataSets/LTE_TNL_BTS_Hourly')
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

## Import & Concat Data Set


```python
# import & concat all the Hourly KPIs
df=pd.concat((pd.read_csv(file,skiprows=range(6),\
    skipfooter=1,engine='python',
    parse_dates=["Date","Time"],na_values=['NIL','/0'],\
                usecols=['Date','Time','eNodeB Name',
                        'Radio Network Unavailability Rate_Cell','L.E-RAB.FailEst.TNL']) for file in filtered_file_list)).\
                        sort_values('Date').reset_index()
```

## Calculate TNL Fail Ex Low Availability Intervals


```python
# filter Required Intervals
df0 = df[df['Radio Network Unavailability Rate_Cell'].eq(0)]
# Calculate Site Level TNL Fail
df1 = df0.groupby(['Date','eNodeB Name'])[['L.E-RAB.FailEst.TNL']].sum().reset_index()
```

## Calculate Low Availability Intervals Count


```python
# Set 'Date' column as the index
df.set_index('Date', inplace=True)

# Calculate intervals
df2 = df.groupby(['Date','eNodeB Name']).apply(lambda x: pd.Series({
        'Total Count of (Down Time)>0 between 0:00-23:00': (x['Radio Network Unavailability Rate_Cell'].gt(0)).sum(),
        'Total Count of (Down Time)>0 between 9:00-21:00': (x.set_index("Time").between_time('9:00', '21:00')['Radio Network Unavailability Rate_Cell'].gt(0)).sum()
    })).reset_index()


# Resetting the index after groupby to get the 'Date' column back
df2.reset_index(drop=True, inplace=True)
```

## Export TNL Fail Ex Low Availability Intervals


```python
# Set Export File Path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/TNL_Ava_Interval/TNL'
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
    output_file = f'{new_max_num:02d}_TNL_EX_Outage_{unique_date_str}.csv'
    
    # Filter the DataFrame for the current unique date
    df_filtered = df1[df1['Date'] == unique_date]
    
    # Extract the time portion from the 'Time' column and convert to 24-hour format
    #df_filtered['Time'] = df_filtered['Time'].dt.strftime('%H:%M')
    
    # Export to CSV
    df_filtered.to_csv(output_file, index=False)

    # Increment the maximum number for the next iteration
    new_max_num += 1
```

## Export Low Availability Intervals Count


```python
# Set Export File Path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/TNL_Ava_Interval/Availability'
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
for unique_date in df2['Date'].unique():
    unique_date_str = unique_date.strftime('%d%m%Y')
    
    # Generate the output file name
    output_file = f'{new_max_num:02d}_LTE_Outage_Intervals_{unique_date_str}.csv'
    
    # Filter the DataFrame for the current unique date
    df_filtered = df2[df2['Date'] == unique_date]
    
    # Extract the time portion from the 'Time' column and convert to 24-hour format
    #df_filtered['Time'] = df_filtered['Time'].dt.strftime('%H:%M')
    
    # Export to CSV
    df_filtered.to_csv(output_file, index=False)

    # Increment the maximum number for the next iteration
    new_max_num += 1
```


```python
#re-set all the variable from the RAM
%reset -f
```
