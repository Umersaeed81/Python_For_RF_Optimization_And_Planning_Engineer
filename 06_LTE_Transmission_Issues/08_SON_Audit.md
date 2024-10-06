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

## Import and Concat RF Export (GLOBALPROCSWITCH)


```python
# get Cell File Path from all the Sub folders
df1 = glob('D:/Advance_Data_Sets/RF_Export/LTE/Cell/*.xlsx', recursive=True)

# concat all the Cell DA Files
df2=pd.concat((pd.read_excel(file,skiprows=[0],sheet_name='GLOBALPROCSWITCH',engine='openpyxl',\
    usecols=['*eNodeB Name',
            'The Switch of X2 setup by SON']) for file in df1))\
            .reset_index(drop=True)
```

## Filter Requird Sites


```python
df3 = df2[df2['The Switch of X2 setup by SON'].ne('ON')].reset_index(drop=True)
```

## Export Output


```python
# set the input file path
os.chdir('D:/Advance_Data_Sets/BSS_EI_Output/TNL_Output')
df3.to_excel('GLOBALPROCSWITCH.xlsx',sheet_name="LTE_eNodeB_SON=OFF",engine='openpyxl',na_rep='-',index=False)
```
