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

## Set Input File Path


```python
# Set Input File PATH
folder_path = 'D:/Advance_Data_Sets/GUL/GUL_Output'
os.chdir(folder_path)
```

## Import UMTS Mapping File


```python
df = pd.read_csv('02_UMTS_Cell_Mapping_wrt_Band.csv',usecols=['Project_Key'])
```

## Import LTE Mappint File


```python
df0 = pd.read_csv('06_LTE_Cell_Mapping_wrt_Band.csv',usecols=['Project_Key'])
```

## Concatenate DataFrame


```python
# Concatenate df and df0, then remove duplicates
df1 = pd.concat([df, df0], ignore_index=True).drop_duplicates(ignore_index=True)
```

## Export Output


```python
folder_path = 'D:/Advance_Data_Sets/GUL/GUL_Output'
os.chdir(folder_path)
df1.to_csv('08_Project_Key.csv',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
