#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

## Import Required Libraries & Set File Path


```python
# Python libraries
import os
from glob import glob
import pandas as pd
```

## Set Input File Path


```python
# Set Working Folder Path
path = 'D:/Advance_Data_Sets/EI_Parser'
os.chdir(path)
```

## Import and concat Excel Files 


```python
# get the file list
filelist= sorted(glob('*.xlsx'))
# get the sheet name in file 0 index
sheets = pd.ExcelFile(filelist[0]).sheet_names

# import and concat Excel Files
df_list = [pd.concat([pd.read_excel(file, header=0, sheet_name=s, na_values=['NIL', '/0'],
                                   dtype={"Cell ID": "str"},
                                   converters={'Integrity': lambda value: '{:,.0f}%'.format(value * 100)})
                                  .assign(Technology=s)
                                  for s in sheets], ignore_index=True)
                                   for file in filelist]

# concat 
df = pd.concat(df_list, ignore_index=True)
```

## Convert the multi rows to single group row


```python
#convert the multi rows to single group row
df0=pd.DataFrame(df.groupby(['Technology','RACaseSerialNo'])['Cell ID'].\
                 apply(';'.join)).reset_index()
```

## Re-shape(pivot_table) data set


```python
# re-shape(pivot_table) data set
df1 = (df0.pivot_table(index=['RACaseSerialNo'], 
                      columns=['Technology'], values='Cell ID', 
                      aggfunc=lambda x: ''.join(str(v) for v in x))).reset_index().fillna('')

```

## Count Cell Count w.r.t Technology


```python
df1['GSM_Cell_Count'] = df1['GSM_IOI_Huawei'].apply(lambda x: set(x.split(';')) if x else '').\
                         apply(lambda x: len(x) if x else 0)


df1['LTE_Cell_Count'] = df1['LTE_UL_Interference_Huawei'].apply(lambda x: set(x.split(';')) if x else '').\
                         apply(lambda x: len(x) if x else 0)


df1['U2100_Cell_Count'] = df1['U2100_Band1_Huawei'].apply(lambda x: set(x.split(';')) if x else '').\
                         apply(lambda x: len(x) if x else 0)


df1['U900_Cell_Count'] = df1['U900_Band8_Huawei'].apply(lambda x: set(x.split(';')) if x else '').\
                         apply(lambda x: len(x) if x else 0)
```

## Export Output


```python
df1.to_excel('huawei_Output.xlsx',index=False)
```

# Compare pre and post

## Import required Libraries


```python
import os
import numpy as np
import pandas as pd
```

## Set Input File Path


```python
# Set Working Folder Path
path = 'D:/Advance_Data_Sets/EI_Parser'
os.chdir(path)
```

## Import Input File Path


```python
g1 = pd.read_csv('Tracker.csv')
g1.head(1)
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>RACaseSerialNo</th>
      <th>Region</th>
      <th>City</th>
      <th>Category</th>
      <th>2G_CI_PTML</th>
      <th>3G_U2100_CI_PTML</th>
      <th>3G_U900_CI_PTML</th>
      <th>4G_CI_PTML</th>
      <th>GSM_IOI_Huawei</th>
      <th>LTE_UL_Interference_Huawei</th>
      <th>U2100_Band1_Huawei</th>
      <th>U900_Band8_Huawei</th>
      <th>GSM_Cell_Count</th>
      <th>LTE_Cell_Count</th>
      <th>U2100_Cell_Count</th>
      <th>U900_Cell_Count</th>
      <th>RSA Status</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>1660-C</td>
      <td>Central</td>
      <td>Lahore</td>
      <td>Pvt. Office</td>
      <td>11699</td>
      <td>NaN</td>
      <td>16946;36946;19383;39383</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>Close</td>
    </tr>
  </tbody>
</table>
</div>



## Convert the values to Sets


```python
g1.loc[:,'GSM_IOI_Huawei']  = g1['GSM_IOI_Huawei'].\
                              replace(' ','').\
                              replace(np.nan,'').str.split(';').apply(set)

g1.loc[:,'U900_Band8_Huawei']  = g1['U900_Band8_Huawei'].\
                              replace(' ','').\
                              replace(np.nan,'').str.split(';').apply(set)

g1.loc[:,'U2100_Band1_Huawei']  = g1['U2100_Band1_Huawei'].\
                              replace(' ','').\
                              replace(np.nan,'').str.split(';').apply(set)

g1.loc[:,'LTE_UL_Interference_Huawei']  = g1['LTE_UL_Interference_Huawei'].\
                              replace(' ','').\
                              replace(np.nan,'').str.split(';').apply(set)
```


```python
g1.loc[:,'2G_CI_PTML']  = g1['2G_CI_PTML'].\
                              replace(' ','').\
                              replace(np.nan,'').str.split(';').apply(set)
g1.loc[:,'3G_U900_CI_PTML']  = g1['3G_U900_CI_PTML'].\
                              replace(' ','').\
                              replace(np.nan,'').str.split(';').apply(set)

g1.loc[:,'3G_U2100_CI_PTML']  = g1['3G_U2100_CI_PTML'].\
                              replace(' ','').\
                              replace(np.nan,'').str.split(';').apply(set)

g1.loc[:,'4G_CI_PTML']  = g1['4G_CI_PTML'].\
                              replace(' ','').\
                              replace(np.nan,'').str.split(';').apply(set)
```

## Compare Sets


```python
g1['GSM_Check']= g1['GSM_IOI_Huawei']==g1['2G_CI_PTML']
g1['U900_Check']= g1['U900_Band8_Huawei']==g1['3G_U900_CI_PTML']
g1['U2100_Check']= g1['U2100_Band1_Huawei']==g1['3G_U2100_CI_PTML']
g1['LTE_Check']= g1['LTE_UL_Interference_Huawei']==g1['4G_CI_PTML']
```

## Output


```python
g1.to_csv('f.csv',index=False)
```
