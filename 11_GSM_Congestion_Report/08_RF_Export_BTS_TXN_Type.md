#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# GSM RF Untilization (NE Data Set)

## Import requied Libraries


```python
import io
import os
import glob
import fnmatch
import zipfile
import pandas as pd
from itertools import permutations
```

## Set File Path


```python
# Set RF Export File Path
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(folder_path)
```

## Import Required File


```python
df = pd.read_csv('04_RF_Export.csv')
```


```python
df0 = pd.read_csv('07_BTS_TXN_Type.csv',usecols=['BSCName','CELLNAME','Site Type','TNX_Type']).\
        rename(columns={'BSCName': 'NE Information;BSCName',\
                   'CELLNAME':'NE Information;CELLNAME',\
                   'Site Type': 'NE Information;Site_Type',\
                    'TNX_Type': 'NE Information;TNX_Type'})
```

## Merge RF Export with BTS_TXN Data Type Date Set


```python
df1 = pd.merge(df,df0,how='left',on=['NE Information;BSCName','NE Information;CELLNAME'])
```


```python
df2 = df1[['NE Information;Region','NE Information;BSCName', 'NE Information;BTSNAME',
       'NE Information;CELLNAME', 'NE Information;LAC', 'NE Information;CI',
       'NE Information;Site_Type', 'NE Information;TNX_Type',
       'NE Information;ACTSTATUS', 'NE Information;Cluster','NE Information;Engineer', 
       'Configuration;GSM(Cell Level)', 'Configuration;DCS(Cell Level)',
       'Configuration;Site Configuration', 'Configuration;GSM(Site Level)',
       'Configuration;DCS(Site Level)', 'Configuration;Total TRX Count',
       'Cell Level Parameters;TCHBUSYTHRES',
       'Cell Level Parameters;AMRTCHHPRIORLOAD',
       'Cell Level Parameters;TCHTRICBUSYOVERLAYTHR',
       'Cell Level Parameters;TCHTRIBUSYUNDERLAYTHR',
       'Cell Level Parameters;LOWRXLEVOLFORBIDSWITCH',
       'Cell Level Parameters;OPTILAYER',
       'Cell Level Parameters;HOALGOPERMLAY',
       'Cell Level Parameters;ACCESSOPTILAY',
       'Cell Level Parameters;OTOURECEIVETH',
       'Cell Level Parameters;UTOORECTH', 'Cell Level Parameters;VAMOSSWITCH',
       'Cell Level Parameters;IDLESDTHRES', 'Cell Level Parameters;CELLMAXSD',
       #'Channel Type;CBCCH', 
           'Channel Type;MBCCH', 'Channel Type;PDTCH',
       'Channel Type;SDCCH8', 'Channel Type;SDCCH_CBCH', 'Channel Type;TCHFR',
       'Channel Type;TCHHR']]
```

## Export


```python
# Set Output Path
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(folder_path)
# Export
df2.to_csv('08_RF_BTS_TNX.csv',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
