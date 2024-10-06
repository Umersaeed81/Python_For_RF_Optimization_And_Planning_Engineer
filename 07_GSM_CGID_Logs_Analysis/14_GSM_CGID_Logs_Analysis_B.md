#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

## TCH Availability Rate(%) (BTS-Hourly Level) - Count Calculation

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
# set File Path
working_directory = 'D:/Advance_Data_Sets/TXN_DataSets/GSM_BTS_Hourly'
os.chdir(working_directory)
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

## Import and Concat BTS Hourly KPIs


```python
# concat all the BTS Hourly Files
df=pd.concat((pd.read_csv(file,skiprows=range(6),\
            skipfooter=1,engine='python',\
            usecols=['Date','Time','GBSC','Site Name','TCH Availability Rate(%)'],
            parse_dates=["Date","Time"],na_values=['NIL','/0']) for file in filtered_file_list)).\
            sort_values('Date').reset_index()
```

## Calculate Intervals


```python
# Calculate intervals
df0 = df.groupby(['Date','GBSC','Site Name']).apply(lambda x: pd.Series({
        'Total Count of (TCH Ava Rate)<100 between 0:00-23:00': (x['TCH Availability Rate(%)'].lt(100)).sum(),
        'Total Count of (TCH Ava Rate)<100 between 9:00-21:00': (x.set_index("Time").between_time('9:00', '21:00')['TCH Availability Rate(%)'].lt(100)).sum()
    })).reset_index()
```

## Export Outage Intervals


```python
# Set Export File Path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/2G_Outage_Intervals_BTS_Level'
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
for unique_date in df0['Date'].unique():
    unique_date_str = unique_date.strftime('%d%m%Y')
    
    # Generate the output file name
    output_file = f'{new_max_num:02d}_2G_Outage_BTS_{unique_date_str}.csv'
    
    # Filter the DataFrame for the current unique date
    df_filtered = df0[df0['Date'] == unique_date]
    
    # Extract the time portion from the 'Time' column and convert to 24-hour format
    #df_filtered['Time'] = df_filtered['Time'].dt.strftime('%H:%M')
    
    # Export to CSV
    df_filtered.to_csv(output_file, index=False)

    # Increment the maximum number for the next iteration
    new_max_num += 1
```


```python
# re-set all the variable from the RAM
%reset -f
```
