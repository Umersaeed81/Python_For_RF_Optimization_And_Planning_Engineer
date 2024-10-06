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
# Python libraries
import os
import zipfile
import pandas as pd
from glob import glob
import warnings
warnings.simplefilter("ignore")
```

## Import and Concat RF Export (Base Station Transport Data)


```python
# get Cell File Path from all the Sub folders
df = glob('D:/Advance_Data_Sets/RF_Export/LTE/Cell/ConfigurationData*.xlsx', recursive=True)
# concat all the Cell Base Station Transport Data Files
df0=pd.concat((pd.read_excel(file,skiprows=[0],\
            sheet_name='Base Station Transport Data',engine='openpyxl') for file in df))\
            .reset_index(drop=True)
```

## Convert MBTS Site IDs and eNodeB Name in short Format


```python
df0.insert(1, 'MBTS_SiteID', df0['*Name'].str.split('_', n=2).str[:2].str.join('_'))
df0.insert(4, 'SiteID', df0['*eNodeB Name'].str.split('_', n=1).str[0])
```

## Export Output


```python
# set the Output file path
os.chdir('D:/Advance_Data_Sets/BSS_EI_Output/TNL_Output')
# Write excel file with default behaviour.
with pd.ExcelWriter("04_LTE_TXN_Availability_Issues.xlsx",date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    df0.to_excel(writer,sheet_name="Base_Station_Transport_Data",engine='openpyxl',na_rep='-',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
