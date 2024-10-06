#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# GSM RF Untilization (BTS & TXN Type Identification)

## Import requied Libraries


```python
import os
import glob
from glob import glob
import zipfile
import pandas as pd
import warnings
warnings.simplefilter("ignore")
```


```python
# set file path
working_directory = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(working_directory)
```

## Import Required Data Sets


```python
df = pd.read_csv('05_BTS_Type.csv')
```


```python
df0 = pd.read_csv('06_TXN_Type.csv').rename(columns={'GBSC':'BSCName','Site Name':'BTSNAME'})
```

## Mege Data Frame


```python
df1 = pd.merge(df,df0,how='left',on=['BSCName','BTSNAME']).fillna('-')
```

## Export Output


```python
# Set Output Path
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(folder_path)
# Export
df1.to_csv('07_BTS_TXN_Type.csv',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
