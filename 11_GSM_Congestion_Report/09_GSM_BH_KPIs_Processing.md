#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# GSM RF Untilization (GSM KPIs Processing)

## Import Required Libraries


```python
import os
import zipfile
import numpy as np
import pandas as pd
from glob import glob

import warnings
warnings.simplefilter("ignore")
```

## Import Cell BH KPIs


```python
# Get all zip files in directory and sort them
df0 = sorted(glob('D:/Advance_Data_Sets/KPIs_Analysis/BH_KPIs_Cell_Level/GSM/*.zip')) 
# Concatenate all files in df0 and keep the last 10 days of data
df = pd.concat((pd.read_csv(file, skiprows=range(6),
                            skipfooter=1, engine='python',
                            usecols=['Date', 'GBSC', 'Cell CI', 'GOS-SDCCH(%)','Cell Name',
                                     'Mobility TCH GOS(%)', 'CallSetup TCH GOS(%)',
                                     'K3015:Available TCHs','Global Traffic','U_EFR TRAFFIC',
                                    'U_AMR HR TRAFFIC','900 TRAFFIC (BSC6900)','1800 TRAFFIC (BSC6900)',
                                     'TCH Availability Rate','Payload(GB)'],
                            parse_dates=["Date"], na_values=['NIL', '/0'],dtype={'Cell CI':str})
                for file in df0)) \
        .sort_values('Date') \
        .set_index(['Date']) \
        .last('14D') \
        .reset_index()

# Round the 'K3015:Available TCHs' column
df['rounded_K3015'] = df['K3015:Available TCHs'].apply(round).astype(str)              # round to nearest integer
```

## Calculate Traffif Shares


```python
# Full Rate & Half Rate Traffic Share
df['FR_Traffic_Share'] = (df['U_EFR TRAFFIC']/df['Global Traffic'])*100
df['HR_Traffic_Share'] = (df['U_AMR HR TRAFFIC']/df['Global Traffic'])*100
# 900 & 1800 Traffic Share
df['900_Traffic_Share'] = (df['900 TRAFFIC (BSC6900)']/df['Global Traffic'])*100
df['1800_Traffic_Share'] = (df['1800 TRAFFIC (BSC6900)']/df['Global Traffic'])*100
```

## Import Erlang B Table


```python
#https://telecom-knowledge.blogspot.com/p/erlang-b.html
# set input File Path
folder_path = 'D:/Advance_Data_Sets/KPIs_Analysis/Inter_BSC_HSR/BSC_Def'
os.chdir(folder_path)
# Read in the erlang_b.txt file and rename columns
df1 = pd.read_excel('erlang_b_table.xlsx',
                 dtype={'rounded_K3015':str})
```

## Merge KPIs and Erlang B Table


```python
df2 = pd.merge(df,df1[['rounded_K3015','Offer_Traffic_Erlang']],on=['rounded_K3015'],how='left')
```

## Calcuate RF Utilization


```python
df2['RF_Utilization'] = (df2['Global Traffic']/df2['Offer_Traffic_Erlang'])*100
```

## Select the requied Columns


```python
df3 = df2[['Date', 'GBSC','Cell Name', 'Cell CI', 'GOS-SDCCH(%)', 'CallSetup TCH GOS(%)',
       'Mobility TCH GOS(%)','Global Traffic','FR_Traffic_Share',
        'HR_Traffic_Share', '900_Traffic_Share', '1800_Traffic_Share','TCH Availability Rate','RF_Utilization','Payload(GB)']]
```

## Calculate Average and Again


```python
df4 = df3.groupby(['GBSC','Cell Name']).apply(lambda x: pd.Series({
    # GOS SDCCH
    'GOS-SDCCH(%);Average': np.average(x['GOS-SDCCH(%)'].dropna()),             
    'GOS-SDCCH(%);Againg>=0.1': (x['GOS-SDCCH(%)'].ge(0.1)).sum(),
    'GOS-SDCCH(%);Max': (x['GOS-SDCCH(%)']).max(),
    'GOS-SDCCH(%);Min': (x['GOS-SDCCH(%)']).min(),
    'GOS-SDCCH(%);3rd Peak value': x['GOS-SDCCH(%)'].dropna().nlargest(3).iloc[-1] if len(x['GOS-SDCCH(%)'].dropna()) >= 3 else np.nan,
    # CS TCH GOS    
    'CallSetup TCH GOS(%);Average': np.average(x['CallSetup TCH GOS(%)'].dropna()),
    'CallSetup TCH GOS(%);Againg>=2': (x['CallSetup TCH GOS(%)'].ge(2)).sum(),
    'CallSetup TCH GOS(%);Max': (x['CallSetup TCH GOS(%)']).max(),
    'CallSetup TCH GOS(%);Min': (x['CallSetup TCH GOS(%)']).min(),
    'CallSetup TCH GOS(%);3rd Peak value': x['CallSetup TCH GOS(%)'].dropna().nlargest(3).iloc[-1] if len(x['CallSetup TCH GOS(%)'].dropna()) >= 3 else np.nan,
    # MoB GOS  
    'Mobility TCH GOS(%);Average': np.average(x['Mobility TCH GOS(%)'].dropna()),
    'Mobility TCH GOS(%);Againg>=2': (x['Mobility TCH GOS(%)'].ge(2)).sum(),
    'Mobility TCH GOS(%);Max': (x['Mobility TCH GOS(%)']).max(),
    'Mobility TCH GOS(%);Min': (x['Mobility TCH GOS(%)']).min(),
    'Mobility TCH GOS(%);3rd Peak value': x['Mobility TCH GOS(%)'].dropna().nlargest(3).iloc[-1] if len(x['Mobility TCH GOS(%)'].dropna()) >= 3 else np.nan,
    # CS Traffic 
    'Global Traffic;Average': np.average(x['Global Traffic'].dropna()),
    'Global Traffic;Max': (x['Global Traffic']).max(),
    'Global Traffic;Min': (x['Global Traffic']).min(),
    'Global Traffic;3rd Peak value': x['Global Traffic'].dropna().nlargest(3).iloc[-1] if len(x['Global Traffic'].dropna()) >= 3 else np.nan,
    # Full Rate Traffic
    'FR_Traffic_Share;Average': np.average(x['FR_Traffic_Share'].dropna()),
    'FR_Traffic_Share;Max': (x['FR_Traffic_Share']).max(),
    'FR_Traffic_Share;Min': (x['FR_Traffic_Share']).min(),
    'FR_Traffic_Share;3rd Peak value': x['FR_Traffic_Share'].dropna().nlargest(3).iloc[-1] if len(x['FR_Traffic_Share'].dropna()) >= 3 else np.nan,
    # Half Rate Traffic
    'HR_Traffic_Share;Average': np.average(x['HR_Traffic_Share'].dropna()),
    'HR_Traffic_Share;Max': (x['HR_Traffic_Share']).max(),
    'HR_Traffic_Share;Min': (x['HR_Traffic_Share']).min(),  
    'HR_Traffic_Share;3rd Peak value': x['HR_Traffic_Share'].dropna().nlargest(3).iloc[-1] if len(x['HR_Traffic_Share'].dropna()) >= 3 else np.nan,
    # 900 Traffic
    '900_Traffic_Share;Average': np.average(x['900_Traffic_Share'].dropna()),
    '900_Traffic_Share;Max': (x['900_Traffic_Share']).max(),
    '900_Traffic_Share;Min': (x['900_Traffic_Share']).min(),
    '900_Traffic_Share;3rd Peak value': x['900_Traffic_Share'].dropna().nlargest(3).iloc[-1] if len(x['900_Traffic_Share'].dropna()) >= 3 else np.nan,
    # 1800 Traffic
    '1800_Traffic_Share;Average': np.average(x['1800_Traffic_Share'].dropna()),
    '1800_Traffic_Share;Max': (x['1800_Traffic_Share']).max(),
    '1800_Traffic_Share;Min': (x['1800_Traffic_Share']).min(),
    '1800_Traffic_Share;3rd Peak value': x['1800_Traffic_Share'].dropna().nlargest(3).iloc[-1] if len(x['1800_Traffic_Share'].dropna()) >= 3 else np.nan,
    # Pay load    
    'Payload;Average': np.average(x['Payload(GB)'].dropna()),
    'Payload;Max': (x['Payload(GB)']).max(),
    'Payload;Min': (x['Payload(GB)']).min(),
    'Payload;3rd Peak value': x['Payload(GB)'].dropna().nlargest(3).iloc[-1] if len(x['Payload(GB)'].dropna()) >= 3 else np.nan,
    # Availability     
    'TCH Availability Rate;Average': np.average(x['TCH Availability Rate'].dropna()),
    'TCH Availability Rate;Againg<100': (x['TCH Availability Rate'].lt(100)).sum(),
    'TCH Availability Rate;Max': (x['TCH Availability Rate']).max(),
    'TCH Availability Rate;Min': (x['TCH Availability Rate']).min(), 
    # RF_Utilization   
    'RF_Utilization;Average': np.average(x['RF_Utilization'].dropna()),
    'RF_Utilization;Max': (x['RF_Utilization']).max(),
    'RF_Utilization;Min': (x['RF_Utilization']).min(),
    'RF_Utilization;3rd Peak value': x['RF_Utilization'].dropna().nlargest(3).iloc[-1] if len(x['RF_Utilization'].dropna()) >= 3 else np.nan
    })).reset_index()
```

## Re-name Header


```python
# Rename the columns of the dataframe
df4 = df4.rename(columns={'GBSC':'NE Information;BSCName',\
                         'Cell Name':'NE Information;CELLNAME'})
```

## RF Export (Processed)


```python
# Set Output Path
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(folder_path)
df5 = pd.read_csv('08_RF_BTS_TNX.csv')
```

## Merge RF Export with KPI


```python
df6 = pd.merge(df5,df4,how='left',on=['NE Information;BSCName','NE Information;CELLNAME']).fillna('-')
```

## Split Header Row


```python
# Split the header
df6.columns = df6.columns.str.split(';', expand=True)
# Reset the index
df6 = df6.reset_index(drop=True)
```

## Export


```python
df6.to_excel('GSM_RF_Utilization.xlsx', sheet_name='RF_Utilization')
```


```python
# re-set all the variable from the RAM
%reset -f
```
