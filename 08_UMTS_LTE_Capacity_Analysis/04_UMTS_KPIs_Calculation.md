#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# UMTS KPIs Calculation

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
                usecols=['Date','RNC','TCP Usage', 'HSDPA Throughput (kbps)',\
               'Total TCP Utilization Rate (Cell-)(%)', 'HSUPA Throughput (kbps)',\
               'HSDPA Users_MAX', 'AMR CS Traffic (Erl)', 'Data Volume (GB)',\
               'NODEBNAME','CELLNAME','VS.TP.UE.0',\
               'VS.TP.UE.1','VS.TP.UE.10.15','VS.TP.UE.16.25',\
               'VS.TP.UE.2','VS.TP.UE.26.35','VS.TP.UE.3',\
               'VS.TP.UE.36.55','VS.TP.UE.4','VS.TP.UE.5',\
               'VS.TP.UE.6.9','VS.TP.UE.More55']) for file in filtered_file_list)).\
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

# Calculate TP Ranges


```python
# TP 0 to 3
df['TP 0-3'] = df['VS.TP.UE.0']+df['VS.TP.UE.1']+df['VS.TP.UE.2']+df['VS.TP.UE.3']
# TP 4 to 5
df['TP 4-5'] = df['VS.TP.UE.4']+df['VS.TP.UE.5']
# TP >36
df['TP >36'] = df['VS.TP.UE.36.55']+df['VS.TP.UE.More55']
```

## Calculate Average UMTS KPIs


```python
# Funcation for drop na values in a columns
def group_wa(series):
    dropped = series.dropna()
    return np.average(dropped)
```


```python
# calculate average of each column after droping na values
df0 = {'AMR CS Traffic (Erl)': [('Average',group_wa)],
            'TCP Usage': [('Average',group_wa)],
            'HSDPA Throughput (kbps)': [('Average',group_wa),\
                            ('<1024 Aging', lambda x: (x.lt(1024)).sum())],
            'Total TCP Utilization Rate (Cell-)(%)': [('Average',group_wa),\
                            ('>85 Aging', lambda x: (x.gt(85)).sum())],
            'HSUPA Throughput (kbps)': [('Average',group_wa)],
            'HSDPA Users_MAX': [('Average',group_wa)],
            'Data Volume (GB)': [('Average',group_wa)],
            'TP 0-3':[('Average',group_wa)],
            'TP 4-5':[('Average',group_wa)],
            'VS.TP.UE.6.9':[('Average',group_wa)],
            'VS.TP.UE.10.15':[('Average',group_wa)],
            'VS.TP.UE.16.25':[('Average',group_wa)],
             'VS.TP.UE.26.35':[('Average',group_wa)],
            'TP >36':[('Average',group_wa)]}
```


```python
# apply series function 
df1 = df.groupby(['CELLNAME']).agg(df0)
```


```python
# Join Level 1 and Level 2 header mapping
df1.columns = df1.columns.map(' '.join)
```


```python
# re-set index
df1=df1.reset_index()
```

## Check UMTS KPIs Thresholds


```python
# Throguhput Categorization
df1['Throguhput Categorization']= np.where(
                (df1['HSDPA Throughput (kbps) Average']>1000),
                '>1Mbps',
                        np.where(
                        (df1['HSDPA Throughput (kbps) Average']>512) & \
                            (df1['HSDPA Throughput (kbps) Average']<=1000),
                        '> 512 kbps & <= 1 Mbps',
                            np.where(
                            (df1['HSDPA Throughput (kbps) Average']>256) & \
                                (df1['HSDPA Throughput (kbps) Average']<=512),
                            '> 256 kbps & <= 512',
                                np.where(
                                (df1['HSDPA Throughput (kbps) Average']<=256),
                                '<= 256 kbps',
                                'Data Missing'))))
```


```python
# Check Average KPIs Thresholds(Total TCP Utilization Rate (Cell-)(%))
df1['Total TCP Utilization Rate (Cell-)(%)>85 Avg KPIs'] = np.where(
            ((df1['Total TCP Utilization Rate (Cell-)(%) Average']>85)),
            'Yes','No')
```


```python
# Check Average KPIs Thresholds(HSDPA Throughput (kbps))
df1['HSDPA Throughput (kbps)<1024 Avg KPIs'] = np.where(
            ((df1['HSDPA Throughput (kbps) Average']<1024)),
            'Yes','No')
```


```python
# Check Average KPIs Congestion
df1['Congestion Status (Avg KPIs)'] = np.where(
            ((df1['Total TCP Utilization Rate (Cell-)(%)>85 Avg KPIs']=='Yes') & \
             (df1['HSDPA Throughput (kbps)<1024 Avg KPIs']=='Yes')),
            'Yes','No')
```

## Check UMTS KPIs Aging Thresholds


```python
# Total TCP Utilization Rate (Cell-)(%)
df1['Total TCP Utilization Rate (Cell-)(%) >85 Aging>10 Days']= np.where(
                (df1['Total TCP Utilization Rate (Cell-)(%) >85 Aging']<=10),
                'No','Yes')
```


```python
# HSDPA Throughput (kbps)
df1['HSDPA Throughput (kbps) <1024 Aging>10 Days']= np.where(
                (df1['HSDPA Throughput (kbps) <1024 Aging']<=10),
                'No','Yes')
```


```python
# Congestion Status (Aging)
df1['Congestion Status (Aging)'] = np.where(
            ((df1['Total TCP Utilization Rate (Cell-)(%) >85 Aging>10 Days']=='Yes') & \
             (df1['HSDPA Throughput (kbps) <1024 Aging>10 Days']=='Yes')),
            'Yes','No')
```

## Final UMTS Congestion Status


```python
df1['Final Congestion'] = np.where(
            ((df1['Congestion Status (Aging)']=='Yes') & \
             (df1['Congestion Status (Avg KPIs)']=='Yes')),
            'Yes','No')
```

## Re-Orderd the UMTS KPIs


```python
df2= df1 [['CELLNAME','AMR CS Traffic (Erl) Average',
                    'TCP Usage Average',
                    'HSDPA Users_MAX Average',
                    'HSUPA Throughput (kbps) Average',
                    'Data Volume (GB) Average',
                    'HSDPA Throughput (kbps) Average',
                    'HSDPA Throughput (kbps)<1024 Avg KPIs',
                    'HSDPA Throughput (kbps) <1024 Aging',
                    'HSDPA Throughput (kbps) <1024 Aging>10 Days',
                    'Total TCP Utilization Rate (Cell-)(%) Average', 
                    'Total TCP Utilization Rate (Cell-)(%)>85 Avg KPIs',
                    'Total TCP Utilization Rate (Cell-)(%) >85 Aging',
                    'Total TCP Utilization Rate (Cell-)(%) >85 Aging>10 Days',
                    'Throguhput Categorization',
                    'Congestion Status (Avg KPIs)',
                    'Congestion Status (Aging)',
                    'Final Congestion',
                    'TP 0-3 Average', 'TP 4-5 Average',
                    'VS.TP.UE.6.9 Average', 
                    'VS.TP.UE.10.15 Average',
                    'VS.TP.UE.16.25 Average', 
                     'VS.TP.UE.26.35 Average', 
                     'TP >36 Average'
                    ]].rename(columns = {'CELLNAME':'Cell Name',
                                        'HSDPA Users_MAX Average':'HSDPA Users-MAX Average',
                                        'VS.TP.UE.6.9 Average':'TP 6-9 Average', 
                                        'VS.TP.UE.10.15 Average':'TP 10-15 Average',
                                        'VS.TP.UE.16.25 Average':'TP 16-25 Average', 
                                        'VS.TP.UE.26.35 Average':'TP 26-35 Average'})
```

## Generate UMTS Cell IDs


```python
# Generate UMTS Cell ID
df2['UMTS_Cell_ID'] = df2['Cell Name'].str.split('_').str[0]
```

## Export Output


```python
# Set output File Folder Path
folder_path = 'D:/Advance_Data_Sets/GUL/GUL_Output'
os.chdir(folder_path)
# Export Output
df2.to_csv('03_UMTS_Cell_Process_KPIs.csv',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
