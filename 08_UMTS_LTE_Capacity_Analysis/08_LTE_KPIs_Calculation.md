#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# LTE KPIs Calculation

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

## Filter LTE KPIs Files


```python
# List all files in the path
file_list = glob('D:/Advance_Data_Sets/GUL/4G/*.zip')

# Calculate the date for 14 days ago
required_date = datetime.now() - timedelta(days=20)
required_date_str = required_date.strftime("%d%m%Y")

# Filter files that match the required date or the two days before
filtered_file_list = [file for file in file_list if any((required_date + timedelta(days=i)).strftime("%d%m%Y") in file for i in range(20))]

if filtered_file_list:
    print("Filtered File List:", filtered_file_list)
else:
    print(f"No file found with the required dates in the last 14 days.")
```

## Import & Concat LTE Cell Busy KPIs


```python
# import & concat all the Cell Busy KPIs
df=pd.concat((pd.read_csv(file,skiprows=range(6),\
    skipfooter=1,engine='python',
    parse_dates=["Date"],na_values=['NIL','/0'],\
                usecols=['Date','eNodeB Name','Cell Name','LocalCell Id','eNodeB Function Name',\
                'DL PRB Avg Utilization','UL PRB Avg Utilization',\
                'DL User Thrp (Mbps)','Data Volume (GB)',\
                 'Average User Number','L.Traffic.ActiveUser.Max',
                'L.RA.TA.UE.Index0','L.RA.TA.UE.Index1','L.RA.TA.UE.Index2',
                'L.RA.TA.UE.Index3','L.RA.TA.UE.Index4','L.RA.TA.UE.Index5',
                'L.RA.TA.UE.Index6','L.RA.TA.UE.Index7','L.RA.TA.UE.Index8',
                'L.RA.TA.UE.Index9','L.RA.TA.UE.Index10','L.RA.TA.UE.Index11']) for file in filtered_file_list)).\
                 sort_values('Date').set_index(['Date']).last('14D').reset_index()
```

## Replace multiple spaces with single space in 'eNodeB Name' & 'CELLNAME' column


```python
# Replace multiple spaces with single space in 'Cell Name' column
df['Cell Name'] = df['Cell Name'].str.replace(r'\s+', ' ', regex=True)
# Replace multiple spaces with single space in 'NodeB Name' column
df['eNodeB Name'] = df['eNodeB Name'].str.replace(r'\s+', ' ', regex=True)
```

## Calculate TA Ranges


```python
# TA 0 to 3
df['TA 0-3'] = df['L.RA.TA.UE.Index0']+\
               df['L.RA.TA.UE.Index1']+\
               df['L.RA.TA.UE.Index2']+\
               df['L.RA.TA.UE.Index3']
```


```python
# TA 4 to 7
df['TA 4-7'] = df['L.RA.TA.UE.Index4']+\
               df['L.RA.TA.UE.Index5']+\
               df['L.RA.TA.UE.Index6']+\
               df['L.RA.TA.UE.Index7']
```


```python
# TA 8 to 11
df['TA 8-11'] = df['L.RA.TA.UE.Index8']+\
                df['L.RA.TA.UE.Index9']+\
                df['L.RA.TA.UE.Index10']+\
                df['L.RA.TA.UE.Index11']
```

### Calculate Average & aging 4G KPIs 


```python
# Funcation for drop na values in a columns
def group_wa(series):
    dropped = series.dropna()
    return np.average(dropped)
```


```python
# calculate average of each column after droping na values
df0 = {'DL PRB Avg Utilization': [('Average',group_wa),\
                                      ('>70 Aging', lambda x: (x.gt(70)).sum())],
           'UL PRB Avg Utilization': [('Average',group_wa)],
            'DL User Thrp (Mbps)': [('Average',group_wa),\
                                    ('<2 Aging',lambda x: (x.lt(2)).sum())],
            'Data Volume (GB)': [('Average',group_wa)],
            'Average User Number':[('Average',group_wa)],
            'L.Traffic.ActiveUser.Max':[('Average',group_wa)],
            'TA 0-3':[('Average',group_wa)],                        
            'TA 4-7':[('Average',group_wa)],
            'TA 8-11':[('Average',group_wa)]}
```


```python
# apply series function 
df1 = df.groupby(['Cell Name']).agg(df0)
```


```python
# Join Level 1 and Level 2 header mapping
df1.columns = df1.columns.map(' '.join)
```


```python
df1=df1.reset_index()
```

## Check LTE KPIs Thresholds


```python
# Throguhput Categorization
df1['Throguhput Categorization']= np.where(
                (df1['DL User Thrp (Mbps) Average']>=2),
                '>2Mbps',
                        np.where(
                        (df1['DL User Thrp (Mbps) Average']>=0) & \
                            (df1['DL User Thrp (Mbps) Average']<0.5),
                        '>= 0 Mbps & < 0.5',
                            np.where(
                            (df1['DL User Thrp (Mbps) Average']>=0.5) & \
                                (df1['DL User Thrp (Mbps) Average']<1),
                            '>= 0.5 Mbps & < 1.0',
                                np.where(
                                (df1['DL User Thrp (Mbps) Average']>=1.0) & \
                                    (df1['DL User Thrp (Mbps) Average']<1.5),
                                '>= 1.0 Mbps & < 1.5',
                                    np.where(
                                    (df1['DL User Thrp (Mbps) Average']>=1.5) & \
                                        (df1['DL User Thrp (Mbps) Average']<2),
                                    '>= 1.5 Mbps & < 2.0',
                                    'Data Missing')))))
```


```python
# Check Average KPIs Thresholds(DL PRB Avg Utilization)
df1['DL PRB Avg Utilization>70 Avg KPIs']= np.where(
                    ((df1['DL PRB Avg Utilization Average']>70)),
                    'Yes','No')
```


```python
# Check Average KPIs Thresholds(DL User Thrp (Mbps))
df1['DL User Thrp (Mbps)<2 Avg KPIs']= np.where(
                    ((df1['DL User Thrp (Mbps) Average']<2)),
                    'Yes','No')
```


```python
# Check Average KPIs Congestion
df1['Congestion Status (Avg KPIs)'] = np.where(
            ((df1['DL PRB Avg Utilization>70 Avg KPIs']=='Yes') & \
             (df1['DL User Thrp (Mbps)<2 Avg KPIs']=='Yes')),
            'Yes','No')
```

## Check LTE KPIs Aging Thresholds


```python
# DL PRB Avg Utilization
df1['DL PRB Avg Utilization >70 Aging>10 Days']= np.where(
                (df1['DL PRB Avg Utilization >70 Aging']<10),
                'No','Yes')
```


```python
# DL User Thrp (Mbps)
df1['DL User Thrp (Mbps) <2 Aging>10 Days']= np.where(
                (df1['DL User Thrp (Mbps) <2 Aging']<10),
                'No','Yes')
```


```python
# Congestion Status (againg)
df1['Congestion Status (againg)'] = np.where(
            ((df1['DL PRB Avg Utilization >70 Aging>10 Days']=='Yes') & \
             (df1['DL User Thrp (Mbps) <2 Aging>10 Days']=='Yes')),
            'Yes','No')
```

## Final LTE Congestion Status


```python
df1['Final Congestion'] = np.where(
            ((df1['Congestion Status (Avg KPIs)']=='Yes') & \
             (df1['Congestion Status (againg)']=='Yes')),
            'Yes','No')
```

## Re-Orderd the LTE KPIs


```python
df2 =  df1 [['Cell Name', 'Data Volume (GB) Average',
                    'Average User Number Average','L.Traffic.ActiveUser.Max Average',
                    'UL PRB Avg Utilization Average',
                    'DL PRB Avg Utilization Average','DL PRB Avg Utilization>70 Avg KPIs',
                    'Throguhput Categorization',
                    'DL PRB Avg Utilization >70 Aging',
                    'DL PRB Avg Utilization >70 Aging>10 Days',
                    'DL User Thrp (Mbps) Average',
                    'DL User Thrp (Mbps)<2 Avg KPIs',
                    'DL User Thrp (Mbps) <2 Aging',
                    'DL User Thrp (Mbps) <2 Aging>10 Days',
                    'Congestion Status (Avg KPIs)',
                    'Congestion Status (againg)',
                    'Final Congestion',
                    'TA 0-3 Average', 'TA 4-7 Average','TA 8-11 Average']]
```

## Generate LTE Cell IDs


```python
# Generate LTE Cell ID
df2['LTE_Cell_ID'] = df2['Cell Name'].str.split('_').str[0]
```

## Export LTE Procssed KPIs


```python
# Set Output Folder Path
folder_path = 'D:/Advance_Data_Sets/GUL/GUL_Output'
os.chdir(folder_path)
# Export Output
df2.to_csv('07_LTE_Cell_Process_KPIs.csv',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
