#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# GSM RF Untilization (Final RF Export)

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
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(folder_path)
```

## Import Input Data Set


```python
df = pd.read_csv('00_RF_Export_Cell_Para.csv')
```


```python
df0 = pd.read_csv('01_RF_Export_Ch_Type.csv')
```


```python
df1 = pd.read_csv('02_RF_Export_Frequency.csv',dtype={'NE Information;CI': str})
```


```python
df2 =  pd.read_csv('03_PRS.csv' ,dtype={'NE Information;CI': str})
```

## Merge Data Frame (RF Export)


```python
df3 = df1.merge(df,on=['NE Information;BSCName','NE Information;CELLNAME'],how='left').\
            merge(df0,on=['NE Information;BSCName','NE Information;CELLNAME'],how='left')
```

## Merge Final RF Export with PRS


```python
df4 = pd.merge(df3,df2,how='left',on=['NE Information;BSCName','NE Information;CI'])
```

## Requird Sequence


```python
df5 = df4[['NE Information;BSCName', 'NE Information;BTSNAME',
       'NE Information;CELLNAME', 'NE Information;LAC', 'NE Information;CI',
       'NE Information;ACTSTATUS','NE Information;Cluster', 
       'NE Information;Engineer', 'NE Information;Region',
        'Configuration;GSM(Cell Level)',
       'Configuration;DCS(Cell Level)', 'Configuration;Site Configuration',
       'Configuration;GSM(Site Level)', 'Configuration;DCS(Site Level)',
       'Configuration;Total TRX Count', 'Cell Level Parameters;TCHBUSYTHRES',
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
        'Channel Type;MBCCH','Channel Type;PDTCH','Channel Type;SDCCH8',
        'Channel Type;SDCCH_CBCH','Channel Type;TCHFR','Channel Type;TCHHR']]
```

## Export 


```python
# Set Output Path
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(folder_path)
df5.to_csv('04_RF_Export.csv',index=False)
```

## Re-Set Variables


```python
# re-set all the variable from the RAM
%reset -f
```
