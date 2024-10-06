#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# BSS & EI Issues (LTE Hourly)

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

## LTE Cell Hourly KPIs File Path


```python
# Set Path for UMTS Cell DA KPIs
path_cell_hourly = 'D:/Advance_Data_Sets/KPIs_Analysis/Hourly_KPIs_Cell_Level/LTE'
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

    Filtered File List: ['LTE_Cell_Hourly_01_11_03102024.zip', 'LTE_Cell_Hourly_12_23_03102024.zip']
    

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
                       usecols=['Date','Time','eNodeB Name','Cell Name','LocalCell Id',\
                        'eNodeB Function Name','Radio Network Unavailability Rate_Cell',\
                        'L.UL.Interference.Avg(dBm)'],
                      parse_dates=["Date","Time"], na_values=['NIL','/0'],
                      dtype={"LocalCell Id" : str,'Radio Network Unavailability Rate_Cell':float,\
                        'L.UL.Interference.Avg(dBm)':float})\
                      for file in filtered_file_list)
```


```python
df0.head()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Date</th>
      <th>Time</th>
      <th>eNodeB Name</th>
      <th>Cell Name</th>
      <th>LocalCell Id</th>
      <th>eNodeB Function Name</th>
      <th>L.UL.Interference.Avg(dBm)</th>
      <th>Radio Network Unavailability Rate_Cell</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>2024-10-03</td>
      <td>2024-10-04</td>
      <td>MBTS_3146_Ghalib Market Gulberg Lahore-Z3 (3G-...</td>
      <td>4G-CI-121575-3_Ghalib Market Gulberg Lahore-Z3...</td>
      <td>3</td>
      <td>4G-CI-121575_Ghalib Market Gulberg Lahore-Z3(3...</td>
      <td>-112.0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>1</th>
      <td>2024-10-03</td>
      <td>2024-10-04</td>
      <td>MBTS_3146_Ghalib Market Gulberg Lahore-Z3 (3G-...</td>
      <td>4G-CI-121575-2_Ghalib Market Gulberg Lahore-Z3...</td>
      <td>2</td>
      <td>4G-CI-121575_Ghalib Market Gulberg Lahore-Z3(3...</td>
      <td>-112.0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>2</th>
      <td>2024-10-03</td>
      <td>2024-10-04</td>
      <td>MBTS_3146_Ghalib Market Gulberg Lahore-Z3 (3G-...</td>
      <td>4G-CI-121575-1_Ghalib Market Gulberg Lahore-Z3...</td>
      <td>1</td>
      <td>4G-CI-121575_Ghalib Market Gulberg Lahore-Z3(3...</td>
      <td>-113.0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>3</th>
      <td>2024-10-03</td>
      <td>2024-10-04</td>
      <td>MBTS_3050_Cricket House Lahore-Z3 (3G-CI-4016)</td>
      <td>4G-CI-124016-3_Cricket House Lahore-Z3(3050)</td>
      <td>3</td>
      <td>4G-CI-124016_Cricket House Lahore-Z3(3050)</td>
      <td>-112.0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>4</th>
      <td>2024-10-03</td>
      <td>2024-10-04</td>
      <td>MBTS_3050_Cricket House Lahore-Z3 (3G-CI-4016)</td>
      <td>4G-CI-124016-2_Cricket House Lahore-Z3(3050)</td>
      <td>2</td>
      <td>4G-CI-124016_Cricket House Lahore-Z3(3050)</td>
      <td>-108.0</td>
      <td>0.0</td>
    </tr>
  </tbody>
</table>
</div>



## Calculate Degraded Intervals


```python
# Calculate intervals
df1 = df0.groupby(['Date','eNodeB Name','Cell Name','LocalCell Id','eNodeB Function Name']).apply(lambda x: pd.Series({
        'Total Count of (Down Time)>0 between 0:00-23:00': (x['Radio Network Unavailability Rate_Cell'].gt(0)).sum(),
        'Total Count of (Down Time)>0 between 9:00-21:00': (x.set_index("Time").between_time('9:00', '21:00')['Radio Network Unavailability Rate_Cell'].gt(0)).sum(),
        'Total_Interval_UL_Interference>-108': (x['L.UL.Interference.Avg(dBm)'].ge(-108)).sum()
        })).reset_index()
```


```python
df1.head()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Date</th>
      <th>eNodeB Name</th>
      <th>Cell Name</th>
      <th>LocalCell Id</th>
      <th>eNodeB Function Name</th>
      <th>Total Count of (Down Time)&gt;0 between 0:00-23:00</th>
      <th>Total Count of (Down Time)&gt;0 between 9:00-21:00</th>
      <th>Total_Interval_UL_Interference&gt;-108</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>2024-10-03</td>
      <td>MBTS_1000_MSC PECHS Karachi-Z4 (3G-S-2000)</td>
      <td>4G-S-142000-1_MSC PECHS (Gold) Karachi-Z4 (1000)</td>
      <td>1</td>
      <td>4G-S-142000_MSC PECHS (Gold) Karachi-Z4 (1000)</td>
      <td>0</td>
      <td>0</td>
      <td>16</td>
    </tr>
    <tr>
      <th>1</th>
      <td>2024-10-03</td>
      <td>MBTS_1000_MSC PECHS Karachi-Z4 (3G-S-2000)</td>
      <td>4G-S-142000-2_MSC PECHS (Gold) Karachi-Z4 (1000)</td>
      <td>2</td>
      <td>4G-S-142000_MSC PECHS (Gold) Karachi-Z4 (1000)</td>
      <td>0</td>
      <td>0</td>
      <td>9</td>
    </tr>
    <tr>
      <th>2</th>
      <td>2024-10-03</td>
      <td>MBTS_1000_MSC PECHS Karachi-Z4 (3G-S-2000)</td>
      <td>4G-S-142000-3_MSC PECHS (Gold) Karachi-Z4 (1000)</td>
      <td>3</td>
      <td>4G-S-142000_MSC PECHS (Gold) Karachi-Z4 (1000)</td>
      <td>0</td>
      <td>0</td>
      <td>17</td>
    </tr>
    <tr>
      <th>3</th>
      <td>2024-10-03</td>
      <td>MBTS_1001_State Life Building (Gold) Karachi-Z...</td>
      <td>4G-S-142001-1_State Life Building (Gold) Karac...</td>
      <td>1</td>
      <td>4G-S-142001_State Life Building (Gold) Karachi...</td>
      <td>0</td>
      <td>0</td>
      <td>12</td>
    </tr>
    <tr>
      <th>4</th>
      <td>2024-10-03</td>
      <td>MBTS_1001_State Life Building (Gold) Karachi-Z...</td>
      <td>4G-S-142001-2_State Life Building (Gold) Karac...</td>
      <td>2</td>
      <td>4G-S-142001_State Life Building (Gold) Karachi...</td>
      <td>0</td>
      <td>0</td>
      <td>4</td>
    </tr>
  </tbody>
</table>
</div>



## Idendtify the Max RTWP for LTE


```python
df0.head(2)
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Date</th>
      <th>Time</th>
      <th>eNodeB Name</th>
      <th>Cell Name</th>
      <th>LocalCell Id</th>
      <th>eNodeB Function Name</th>
      <th>L.UL.Interference.Avg(dBm)</th>
      <th>Radio Network Unavailability Rate_Cell</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>2024-10-03</td>
      <td>2024-10-04</td>
      <td>MBTS_3146_Ghalib Market Gulberg Lahore-Z3 (3G-...</td>
      <td>4G-CI-121575-3_Ghalib Market Gulberg Lahore-Z3...</td>
      <td>3</td>
      <td>4G-CI-121575_Ghalib Market Gulberg Lahore-Z3(3...</td>
      <td>-112.0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>1</th>
      <td>2024-10-03</td>
      <td>2024-10-04</td>
      <td>MBTS_3146_Ghalib Market Gulberg Lahore-Z3 (3G-...</td>
      <td>4G-CI-121575-2_Ghalib Market Gulberg Lahore-Z3...</td>
      <td>2</td>
      <td>4G-CI-121575_Ghalib Market Gulberg Lahore-Z3(3...</td>
      <td>-112.0</td>
      <td>0.0</td>
    </tr>
  </tbody>
</table>
</div>




```python
df2 = df0[['Date','Time','eNodeB Name','Cell Name','LocalCell Id','eNodeB Function Name','L.UL.Interference.Avg(dBm)']].dropna().reset_index(drop=True)
```


```python
df2.head(2)
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Date</th>
      <th>Time</th>
      <th>eNodeB Name</th>
      <th>Cell Name</th>
      <th>LocalCell Id</th>
      <th>eNodeB Function Name</th>
      <th>L.UL.Interference.Avg(dBm)</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>2024-10-03</td>
      <td>2024-10-04</td>
      <td>MBTS_3146_Ghalib Market Gulberg Lahore-Z3 (3G-...</td>
      <td>4G-CI-121575-3_Ghalib Market Gulberg Lahore-Z3...</td>
      <td>3</td>
      <td>4G-CI-121575_Ghalib Market Gulberg Lahore-Z3(3...</td>
      <td>-112.0</td>
    </tr>
    <tr>
      <th>1</th>
      <td>2024-10-03</td>
      <td>2024-10-04</td>
      <td>MBTS_3146_Ghalib Market Gulberg Lahore-Z3 (3G-...</td>
      <td>4G-CI-121575-2_Ghalib Market Gulberg Lahore-Z3...</td>
      <td>2</td>
      <td>4G-CI-121575_Ghalib Market Gulberg Lahore-Z3(3...</td>
      <td>-112.0</td>
    </tr>
  </tbody>
</table>
</div>




```python
df3 = df2.loc[df2.groupby(['Date','eNodeB Name','Cell Name','LocalCell Id','eNodeB Function Name'])\
            ['L.UL.Interference.Avg(dBm)'].idxmax()].reset_index(drop=True)
```

## Export Max RTWP KPIs (LTE)


```python
# Set Export File Path
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/4G_EI_Max'
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
    output_file = f'{new_max_num:02d}_LTE_RTWP_Max_{unique_date_str}.csv'
    
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
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/4G_EI_Ava_Intervals'
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
    output_file = f'{new_max_num:02d}_LTE_RTWP_DownTime_Intervals_{unique_date_str}.csv'
    
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


```python

```


```python

```
