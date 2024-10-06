#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# GSM RF Untilization (Frequency Export)

## Import requied Libraries


```python
import io
import os
import glob
import fnmatch
import zipfile
import numpy as np
import pandas as pd
from itertools import permutations
```

## Set File Path


```python
# Set RF Export File Path
folder_path = 'D:/Advance_Data_Sets/RF_Export/GSM/00_Frequency_Export'
os.chdir(folder_path)
```

## Unzip File


```python
for file in os.listdir(folder_path):   # get the list of files
    if zipfile.is_zipfile(file): # if it is a zipfile, extract it
        with zipfile.ZipFile(file) as zip_ref: # treat the file as a zip
            folder_name = os.path.splitext(file)[0]  # get the name of the folder (remove '.zip' extension)
            os.makedirs(folder_name, exist_ok=True)  # create the folder if it doesn't exist
            zip_ref.extractall(folder_name)  # extract the contents into the folder
```

## GCELL


```python
# Get GCELL Path from all the Sub folders
df = glob.glob('D:/Advance_Data_Sets/RF_Export/GSM/00_Frequency_Export/**/GCELL.txt', recursive=True) 
# import and concat GCELL File
df0=pd.concat((pd.read_csv(file,header=0,skiprows=[1],
                engine='python',encoding='unicode_escape',\
                usecols=['BSCName','BTSNAME','CELLNAME',
                'LAC','CI','ACTSTATUS']) for file in df)).\
                drop_duplicates().reset_index(drop=True)
```

## GTRX


```python
# Get GTRX Path from all the Sub folders
df1 = glob.glob('D:/Advance_Data_Sets/RF_Export/GSM/00_Frequency_Export/**/GTRX.txt', recursive=True) 
# import and concat GCELLCHMGAD File
df2=pd.concat((pd.read_csv(file,header=0,skiprows=[1],
                engine='python',encoding='unicode_escape',\
                usecols=['BSCName','CELLNAME',
                        'FREQ','ISMAINBCCH']) for file in df1)).\
                drop_duplicates().reset_index(drop=True)
```

## Band Identification


```python
# Band Identification
df2['Band'] = np.where(
            ((df2['FREQ']>=25) & (df2['FREQ']<=62)),
            'PTML-GSM-TRX', 
            np.where(
                    (df2['FREQ']>=512) & (df2['FREQ']<=586), 
                    'PTML-DCS-TRX', 
                     'Other-Band-TRX'))
```

## TRX Count (Band wise-Cell Level)


```python
# Band wise TRX Count (Per Cell)
df3=pd.crosstab([df2["BSCName"],df2["CELLNAME"]],\
                     df2['Band']).reset_index().fillna(0)
```


```python
# Rename the columns of the dataframe
df3 = df3.rename(columns={'PTML-GSM-TRX':'GSM(Cell Level)',\
                         'PTML-DCS-TRX':'DCS(Cell Level)'})
```

## Merge GCELL and TRX Count (Band wise- Cell Leve) DataFrame


```python
df4 = pd.merge(df0,df3,on=['BSCName','CELLNAME'],how='left')
```

## Required Sequence


```python
df4 = df4[['BSCName','BTSNAME','CELLNAME','LAC','CI','ACTSTATUS','GSM(Cell Level)','DCS(Cell Level)']]
```

## Sort DataFrame


```python
# sort Data Frame
df5=df4.sort_values(['BSCName', 'BTSNAME', 'CELLNAME'],ascending=[True,True,True]).reset_index(drop=True)
```

# Change data type of multiple Columns


```python
# Change data type
df5[['GSM(Cell Level)', 'DCS(Cell Level)']]=df5[['GSM(Cell Level)', 'DCS(Cell Level)']].apply(lambda x: x.astype(str))
```

## Site Level Configration


```python
# Total trx per site
df6=pd.merge(df5.groupby(['BSCName', 'BTSNAME'])['GSM(Cell Level)'].apply(';'.join).reset_index()\
             ,df5.groupby(['BSCName', 'BTSNAME'])['DCS(Cell Level)'].apply(';'.join).reset_index()\
             ,on=['BSCName', 'BTSNAME'])
```

## Concat two columns


```python
df6['Site Configuration']=df6['GSM(Cell Level)']+"_"+df6['DCS(Cell Level)']
```

## Caclculate TRX Count(Site Level- Band Wise)


```python
def convert_and_sum(column_name):
    # Convert the string values to lists of integers and store in a new column
    df6[f'{column_name} (Site Level)'] = df6[column_name].str.split(';').apply(lambda x: sum([int(i) for i in x]))


# Apply the function to 'GSM(Cell Level)'
convert_and_sum('GSM(Cell Level)')
# Apply the function to 'DCS(Cell Level)'
convert_and_sum('DCS(Cell Level)')

# Rename the columns of the dataframe
df6 = df6.rename(columns={'GSM(Cell Level) (Site Level)':'GSM(Site Level)',\
                         'DCS(Cell Level) (Site Level)':'DCS(Site Level)'})
```

## Calculate Total TRX


```python
# Function to calculate sum
def calculate_sum(value):
    # Split by semicolon and underscore, then convert to integers
    values = [int(x) for x in value.replace('_', ';').split(';')]
    return sum(values)

# Apply the function to the column
df6['Total TRX Count'] = df6['Site Configuration'].apply(calculate_sum)
```

## Select Required Columns


```python
df6 = df6[['BSCName', 'BTSNAME','Site Configuration', 'GSM(Site Level)', 'DCS(Site Level)',
       'Total TRX Count']]
```

## Merge df5 with df6


```python
df7 = pd.merge(df5,df6,how='left',on=['BSCName','BTSNAME'])
```

## Re-name Header name


```python
# Rename the columns of the dataframe
df7 = df7.rename(columns={'BSCName':'NE Information;BSCName',\
                         'BTSNAME':'NE Information;BTSNAME',\
                          'CELLNAME':'NE Information;CELLNAME',
                          'LAC':'NE Information;LAC',
                          'CI':'NE Information;CI',
                          'ACTSTATUS':'NE Information;ACTSTATUS',
                         'GSM(Cell Level)':'Configuration;GSM(Cell Level)',
                          'DCS(Cell Level)':'Configuration;DCS(Cell Level)',
                          'Site Configuration':'Configuration;Site Configuration',
                          'GSM(Site Level)':'Configuration;GSM(Site Level)',
                          'DCS(Site Level)':'Configuration;DCS(Site Level)',
                          'Total TRX Count':'Configuration;Total TRX Count'
                         })
```

## Export


```python
# Set Output Path
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(folder_path)
df7.to_csv('02_RF_Export_Frequency.csv',index=False)
```

## Re-Set Variables


```python
# re-set all the variable from the RAM
%reset -f
```
