#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore

# GSM RF Untilization (PRS Cluster Defination)

## Import Required Libraries


```python
import io
import os
import glob
import fnmatch
import zipfile
import pandas as pd
from glob import glob
from itertools import permutations
```

## Set Path


```python
working_directory = 'D:/Advance_Data_Sets/PRS/GSM'
os.chdir(working_directory)
```

## UNZIP FILE (Extract Here)


```python
for file in os.listdir(working_directory):
    if zipfile.is_zipfile(os.path.join(working_directory, file)):
        with zipfile.ZipFile(os.path.join(working_directory, file), 'r') as zip_ref:
            zip_ref.extractall(working_directory)
```

## Import Index File


```python
## Get the File Location from Sub folder
df= glob('**/index.xlsx')
# import the Index File
df0 = pd.read_excel(df[0])
```

## Filter Required Cell Groups


```python
# Define a list of categorical values to filter
df1 = ['ABBOTTABAD_CLUSTER_01_Rural',
                        'ABBOTTABAD_CLUSTER_01_Urban',
                        'AHMEDPUREAST_CLUSTER_01_Rural',
                        'AHMEDPUREAST_CLUSTER_01_Urban',
                        'AJK_CLUSTER_01_Rural',
                        'AJK_CLUSTER_01_Urban',
                        'ALIPUR_CLUSTER_01_Rural',
                        'ALIPUR_CLUSTER_01_Urban',
                        'BAHAWALPUR_CLUSTER_01_Rural',
                        'BAHAWALPUR_CLUSTER_01_Urban',
                        'BAHAWALPUR_CLUSTER_02_Rural',
                        'BANNU_CLUSTER_01_Rural',
                        'CHAKWAL_CLUSTER_01_Rural',
                        'CHAKWAL_CLUSTER_01_Urban',
                        'CHAMAN_CLUSTER_21_Rural',
                        'CHAMAN_CLUSTER_21_Urban',
                        'DADU_CLUSTER_15_Rural',
                        'DADU_CLUSTER_15_Urban',
                        'DG_KHAN_CLUSTER_01_Rural',
                        'DG_KHAN_CLUSTER_02_Rural',
                        'DG_KHAN_CLUSTER_02_Urban',
                        'DG_KHAN_CLUSTER_03_Rural',
                        'DG_KHAN_CLUSTER_03_Urban',
                        'DG_KHAN_CLUSTER_04_Rural',
                        'DG_KHAN_CLUSTER_04_Urban',
                        'DI_KHAN_CLUSTER_01_Rural',
                        'DI_KHAN_CLUSTER_01_Urban',
                        'DI_KHAN_CLUSTER_02_Rural',
                        'DI_KHAN_CLUSTER_02_Urban',
                        'DI_KHAN_CLUSTER_03_Rural',
                        'FAISALABAD_CLUSTER_01_Rural',
                        'FAISALABAD_CLUSTER_02_Rural',
                        'FAISALABAD_CLUSTER_03_Rural',
                        'FAISALABAD_CLUSTER_04_Rural',
                        'FAISALABAD_CLUSTER_04_Urban',
                        'FAISALABAD_CLUSTER_05_Rural',
                        'FAISALABAD_CLUSTER_05_Urban',
                        'FAISALABAD_CLUSTER_06_Rural',
                        'FAISALABAD_CLUSTER_06_Urban',
                        'FANA_CLUSTER_01_Rural',
                        'GAWADAR_CLUSTER_20_Rural',
                        'GAWADAR_CLUSTER_20_Urban',
                        'GHOTKI_CLUSTER_09_Rural',
                        'GHOTKI_CLUSTER_09_Urban',
                        'GTROAD_CLUSTER_01_Rural',
                        'GTROAD_CLUSTER_01_Urban',
                        'GUJRANWALA_CLUSTER_01_Rural',
                        'GUJRANWALA_CLUSTER_01_Urban',
                        'GUJRANWALA_CLUSTER_02_Rural',
                        'GUJRANWALA_CLUSTER_02_Urban',
                        'GUJRANWALA_CLUSTER_03_Rural',
                        'GUJRANWALA_CLUSTER_03_Urban',
                        'GUJRANWALA_CLUSTER_04_Rural',
                        'GUJRANWALA_CLUSTER_04_Urban',
                        'GUJRANWALA_CLUSTER_05_Rural',
                        'GUJRANWALA_CLUSTER_05_Urban',
                        'GUJRANWALA_CLUSTER_06_Rural',
                        'HYDERABAD_CLUSTER_01_RURAL',
                        'HYDERABAD_CLUSTER_01_URBAN',
                        'HYDERABAD_CLUSTER_02_RURAL',
                        'HYDERABAD_CLUSTER_02_URBAN',
                        'HYDERABAD_CLUSTER_03_RURAL',
                        'HYDERABAD_CLUSTER_03_URBAN',
                        'HYDERABAD_CLUSTER_04_RURAL',
                        'HYDERABAD_CLUSTER_04_URBAN',
                        'ISLAMABAD_CLUSTER_01_Rural',
                        'ISLAMABAD_CLUSTER_01_Urban',
                        'ISLAMABAD_CLUSTER_02_Rural',
                        'ISLAMABAD_CLUSTER_02_Urban',
                        'ISLAMABAD_CLUSTER_03_Rural',
                        'ISLAMABAD_CLUSTER_03_Urban',
                        'ISLAMABAD_CLUSTER_04_Rural',
                        'ISLAMABAD_CLUSTER_04_Urban',
                        'ISLAMABAD_CLUSTER_05_Rural',
                        'ISLAMABAD_CLUSTER_05_Urban',
                        'JACOBABAD_CLUSTER_12_Rural',
                        'JACOBABAD_CLUSTER_12_Urban',
                        'JAMPUR_CLUSTER_01_Rural',
                        'JAMPUR_CLUSTER_01_Urban',
                        'JHELUM_CLUSTER_01_Rural',
                        'JHELUM_CLUSTER_01_Urban',
                        'JHUNG_CLUSTER_01_Rural',
                        'JHUNG_CLUSTER_01_Urban',
                        'JHUNG_CLUSTER_02_Rural',
                        'JHUNG_CLUSTER_02_Urban',
                        'JHUNG_CLUSTER_03_Rural',
                        'JHUNG_CLUSTER_03_Urban',
                        'JHUNG_CLUSTER_04_Rural',
                        'JHUNG_CLUSTER_04_Urban',
                        'JHUNG_CLUSTER_05_Rural',
                        'JHUNG_CLUSTER_05_Urban',
                        'KARACHI_CLUSTER_01_RURAL',
                        'KARACHI_CLUSTER_01_URBAN',
                        'KARACHI_CLUSTER_02_RURAL',
                        'KARACHI_CLUSTER_02_URBAN',
                        'KARACHI_CLUSTER_03_URBAN',
                        'KARACHI_CLUSTER_04_URBAN',
                        'KARACHI_CLUSTER_05_RURAL',
                        'KARACHI_CLUSTER_05_URBAN',
                        'KARACHI_CLUSTER_06_URBAN',
                        'KARACHI_CLUSTER_07_RURAL',
                        'KARACHI_CLUSTER_07_URBAN',
                        'KARACHI_CLUSTER_08_URBAN',
                        'KARACHI_CLUSTER_09_URBAN',
                        'KARACHI_CLUSTER_10_RURAL',
                        'KARACHI_CLUSTER_10_URBAN',
                        'KARACHI_CLUSTER_11_URBAN',
                        'KARACHI_CLUSTER_12_URBAN',
                        'KARACHI_CLUSTER_13_URBAN',
                        'KARACHI_CLUSTER_14_RURAL',
                        'KARACHI_CLUSTER_14_URBAN',
                        'KARACHI_CLUSTER_15_URBAN',
                        'KARACHI_CLUSTER_16_URBAN',
                        'KARACHI_CLUSTER_17_URBAN',
                        'KARACHI_CLUSTER_18_URBAN',
                        'KARACHI_CLUSTER_19_URBAN',
                        'KARACHI_CLUSTER_20_URBAN',
                        'KARACHI_CLUSTER_21_RURAL',
                        'KARACHI_CLUSTER_21_URBAN',
                        'KARACHI_CLUSTER_22_RURAL',
                        'KARACHI_CLUSTER_22_URBAN',
                        'KASUR_CLUSTER_01_Rural',
                        'KASUR_CLUSTER_02_Rural',
                        'KASUR_CLUSTER_03_Rural',
                        'KASUR_CLUSTER_03_Urban',
                        'KHANPUR_CLUSTER_01_Rural',
                        'KHANPUR_CLUSTER_01_Urban',
                        'KHUZDAR_CLUSTER_17_Rural',
                        'KHUZDAR_CLUSTER_17_Urban',
                        'KOHAT_CLUSTER_01_Rural',
                        'KOHAT_CLUSTER_01_Urban',
                        'KOHAT_CLUSTER_02_Rural',
                        'LAHORE_CLUSTER_01_Rural',
                        'LAHORE_CLUSTER_01_Urban',
                        'LAHORE_CLUSTER_02_Rural',
                        'LAHORE_CLUSTER_02_Urban',
                        'LAHORE_CLUSTER_03_Rural',
                        'LAHORE_CLUSTER_03_Urban',
                        'LAHORE_CLUSTER_04_Urban',
                        'LAHORE_CLUSTER_05_Rural',
                        'LAHORE_CLUSTER_05_Urban',
                        'LAHORE_CLUSTER_06_Rural',
                        'LAHORE_CLUSTER_06_Urban',
                        'LAHORE_CLUSTER_07_Rural',
                        'LAHORE_CLUSTER_07_Urban',
                        'LAHORE_CLUSTER_08_Rural',
                        'LAHORE_CLUSTER_08_Urban',
                        'LAHORE_CLUSTER_09_Rural',
                        'LAHORE_CLUSTER_09_Urban',
                        'LAHORE_CLUSTER_10_Urban',
                        'LAHORE_CLUSTER_11_Rural',
                        'LAHORE_CLUSTER_11_Urban',
                        'LAHORE_CLUSTER_12_Urban',
                        'LAHORE_CLUSTER_13_Urban',
                        'LAHORE_CLUSTER_14_Urban',
                        'LARKANA_CLUSTER_13_Rural',
                        'LARKANA_CLUSTER_13_Urban',
                        'MANSEHRA_CLUSTER_01_Rural',
                        'MANSEHRA_CLUSTER_01_Urban',
                        'MARDAN_CLUSTER_01_Rural',
                        'MARDAN_CLUSTER_01_Urban',
                        'MIRPUR_CLUSTER_01_Rural',
                        'MIRPUR_CLUSTER_01_Urban',
                        'MIRPURKHAS_CLUSTER_05_Rural',
                        'MIRPURKHAS_CLUSTER_05_Urban',
                        'MITHI_CLUSTER_04_Rural',
                        'MITHI_CLUSTER_04_Urban',
                        'MORO_CLUSTER_02_Rural',
                        'MORO_CLUSTER_02_Urban',
                        'MOTORWAY_CLUSTER_01_Rural',
                        'MULTAN_CLUSTER_01_Rural',
                        'MULTAN_CLUSTER_01_Urban',
                        'MULTAN_CLUSTER_02_Rural',
                        'MULTAN_CLUSTER_02_Urban',
                        'MULTAN_CLUSTER_03_Rural',
                        'MULTAN_CLUSTER_03_Urban',
                        'MURREE_CLUSTER_01_Rural',
                        'MURREE_CLUSTER_01_Urban',
                        'MUZAFFARABAD_CLUSTER_01_Rural',
                        'MUZAFFARABAD_CLUSTER_01_Urban',
                        'NAWABSHAH_CLUSTER_01_Rural',
                        'NAWABSHAH_CLUSTER_01_Urban',
                        'NOWSHEHRA_CLUSTER_01_Rural',
                        'NOWSHEHRA_CLUSTER_01_Urban',
                        'PESHAWAR_CLUSTER_01_Rural',
                        'PESHAWAR_CLUSTER_01_Urban',
                        'PESHAWAR_CLUSTER_02_Rural',
                        'PESHAWAR_CLUSTER_02_Urban',
                        'PESHAWAR_CLUSTER_03_Rural',
                        'PESHAWAR_CLUSTER_03_Urban',
                        'PESHAWAR_CLUSTER_04_Rural',
                        'PESHAWAR_CLUSTER_04_Urban',
                        'QUETTAEAST_CLUSTER_19_Urban',
                        'QUETTAWEST_CLUSTER_18_Rural',
                        'QUETTAWEST_CLUSTER_18_Urban',
                        'RAHIMYARKHAN_CLUSTER_01_Rural',
                        'RAHIMYARKHAN_CLUSTER_01_Urban',
                        'RAJANPUR_CLUSTER_01_Rural',
                        'RAJANPUR_CLUSTER_01_Urban',
                        'RAWALPINDI_CLUSTER_01_Rural',
                        'RAWALPINDI_CLUSTER_01_Urban',
                        'RAWALPINDI_CLUSTER_02_Rural',
                        'RAWALPINDI_CLUSTER_02_Urban',
                        'RAWALPINDI_CLUSTER_03_Urban',
                        'RAWALPINDI_CLUSTER_04_Rural',
                        'RAWALPINDI_CLUSTER_04_Urban',
                        'RAWALPINDI_CLUSTER_05_Rural',
                        'RAWALPINDI_CLUSTER_05_Urban',
                        'SADIQABAD_CLUSTER_01_Rural',
                        'SAHIWAL_CLUSTER_01_Rural',
                        'SAHIWAL_CLUSTER_01_Urban',
                        'SAHIWAL_CLUSTER_02_Rural',
                        'SAHIWAL_CLUSTER_02_Urban',
                        'SAHIWAL_CLUSTER_03_Rural',
                        'SAHIWAL_CLUSTER_03_Urban',
                        'SAHIWAL_CLUSTER_04_Rural',
                        'SAHIWAL_CLUSTER_04_Urban',
                        'SAHIWAL_CLUSTER_05_Rural',
                        'SAHIWAL_CLUSTER_05_Urban',
                        'SAHIWAL_CLUSTER_06_Rural',
                        'SAHIWAL_CLUSTER_06_Urban',
                        'SAHIWAL_CLUSTER_07_Rural',
                        'SAHIWAL_CLUSTER_07_Urban',
                        'SAKRAND_CLUSTER_06_Rural',
                        'SAKRAND_CLUSTER_06_Urban',
                        'SHAHDADKOT_CLUSTER_14_Rural',
                        'SHAHDADKOT_CLUSTER_14_Urban',
                        'SIALKOT_CLUSTER_01_Rural',
                        'SIALKOT_CLUSTER_01_Urban',
                        'SIALKOT_CLUSTER_02_Rural',
                        'SIALKOT_CLUSTER_02_Urban',
                        'SIALKOT_CLUSTER_03_Rural',
                        'SIALKOT_CLUSTER_03_Urban',
                        'SIALKOT_CLUSTER_04_Rural',
                        'SIALKOT_CLUSTER_05_Rural',
                        'SIALKOT_CLUSTER_05_Urban',
                        'SIALKOT_CLUSTER_06_Rural',
                        'SIALKOT_CLUSTER_06_Urban',
                        'SIALKOT_CLUSTER_07_Rural',
                        'SIALKOT_CLUSTER_07_Urban',
                        'SIBBI_CLUSTER_16_Rural',
                        'SIBBI_CLUSTER_16_Urban',
                        'SUI_CLUSTER_11_Rural',
                        'SUI_CLUSTER_11_Urban',
                        'SUKKUR_CLUSTER_10_Rural',
                        'SUKKUR_CLUSTER_10_Urban',
                        'SWABI_CLUSTER_01_Rural',
                        'SWABI_CLUSTER_01_Urban',
                        'SWAT_CLUSTER_01_Rural',
                        'SWAT_CLUSTER_01_Urban',
                        'SWAT_CLUSTER_02_Rural',
                        'TALAGANG_CLUSTER_01_Rural',
                        'TAXILA_CLUSTER_01_Rural',
                        'TAXILA_CLUSTER_01_Urban',
                        'THATTA_CLUSTER_07_Rural',
                        'THATTA_CLUSTER_07_Urban',
                        'UMERKOT_CLUSTER_03_Rural',
                        'UMERKOT_CLUSTER_03_Urban' ]

# Filter the DataFrame based on the list of categorical values
df2 = df0[df0['Name(*)'].isin(df1)].reset_index(drop=True)
```

## Get Unique File Name


```python
df3= df2['File Name'].unique()
```


```python
# Add file extensions to the unique_file_names
df4 = [name + '.xlsx' for name in df3]
```

## Get the File Location from Sub folder


```python
df2= glob('**/*.xlsx')
```

## Filter Required Files


```python
# Filter files in df2 based on unique_file_names_with_ext
df5 = [file for file in df2 if any(name in file for name in df4)]
```

## Import Requird Cell Group Files


```python
from datetime import datetime
start_time = datetime.now()
print(start_time)
```

    2024-10-04 10:49:43.630258
    


```python
import dask.dataframe as dd
# Convert df1 to a set
df1_set = set(df1)

# Define a function to read and filter Excel files
def read_and_filter(file):
    dft = pd.read_excel(file, skiprows=range(2), dtype={'Name(*)': str, 'GBSC(*)': str, 'Cell CI': str},
                       na_values=['NIL','/0'], sheet_name='2GBSC@GCell', usecols=['Name(*)','GBSC(*)','Cell CI'])
    filtered_df = dft[dft['Name(*)'].isin(df1_set)]  # Use df1_set instead of df1
    return filtered_df

# Use Dask for parallel processing
ddf = dd.concat([dd.from_pandas(read_and_filter(file), npartitions=4) for file in df5]).reset_index(drop=True)

# Compute the Dask DataFrame to get the final result
final_df = ddf.compute()
```


```python
end_time = datetime.now()
print(end_time)
```

    2024-10-04 11:42:12.021439
    


```python
final_df.head()
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
      <th>Name(*)</th>
      <th>GBSC(*)</th>
      <th>Cell CI</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>31516</td>
    </tr>
    <tr>
      <th>1</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>33498</td>
    </tr>
    <tr>
      <th>2</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>21516</td>
    </tr>
    <tr>
      <th>3</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>23498</td>
    </tr>
    <tr>
      <th>4</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>11516</td>
    </tr>
  </tbody>
</table>
</div>



## Re-name Column Name


```python
# Rename the columns of the dataframe
final_df = final_df.rename(columns={'GBSC(*)': 'NE Information;BSCName',\
                   'Cell CI':'NE Information;CI',\
                      'Name(*)': 'NE Information;Cluster'})
```


```python
final_df.head(2)
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
      <th>NE Information;Cluster</th>
      <th>NE Information;BSCName</th>
      <th>NE Information;CI</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>31516</td>
    </tr>
    <tr>
      <th>1</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>33498</td>
    </tr>
  </tbody>
</table>
</div>



## Cluster Owners


```python
data = {
    # North Clusters
    'ABBOTTABAD_CLUSTER_01_Rural': [('00627299', 'North')],
    'ABBOTTABAD_CLUSTER_01_Urban': [('00627299', 'North')],
    'AJK_CLUSTER_01_Rural': [('00627299', 'North')],
    'AJK_CLUSTER_01_Urban': [('00627299', 'North')],
    'BANNU_CLUSTER_01_Rural': [('00627299', 'North')],
    'CHAKWAL_CLUSTER_01_Rural': [('00627299', 'North')],
    'CHAKWAL_CLUSTER_01_Urban': [('00627299', 'North')],
    'FANA_CLUSTER_01_Rural': [('00627299', 'North')],
    'GTROAD_CLUSTER_01_Rural': [('00627299', 'North')],
    'GTROAD_CLUSTER_01_Urban': [('00627299', 'North')], 
    'ISLAMABAD_CLUSTER_01_Rural': [('00627299', 'North')],
    'ISLAMABAD_CLUSTER_01_Urban': [('00627299', 'North')],
    'ISLAMABAD_CLUSTER_02_Rural': [('00627299', 'North')],
    'ISLAMABAD_CLUSTER_02_Urban': [('00627299', 'North')],
    'ISLAMABAD_CLUSTER_03_Rural': [('00627299', 'North')],
    'ISLAMABAD_CLUSTER_03_Urban': [('00627299', 'North')],
    'ISLAMABAD_CLUSTER_04_Rural': [('00627299', 'North')],
    'ISLAMABAD_CLUSTER_04_Urban': [('00627299', 'North')],
    'ISLAMABAD_CLUSTER_05_Rural': [('00627299', 'North')],
    'ISLAMABAD_CLUSTER_05_Urban': [('00627299', 'North')],
    'JHELUM_CLUSTER_01_Rural': [('00627299', 'North')],
    'JHELUM_CLUSTER_01_Urban': [('00627299', 'North')],
    'KOHAT_CLUSTER_01_Rural': [('00627299', 'North')],
    'KOHAT_CLUSTER_01_Urban': [('00627299', 'North')],
    'KOHAT_CLUSTER_02_Rural': [('00627299', 'North')],
    'MANSEHRA_CLUSTER_01_Rural': [('00627299', 'North')],
    'MANSEHRA_CLUSTER_01_Urban': [('00627299', 'North')],
    'MARDAN_CLUSTER_01_Rural': [('00627299', 'North')],
    'MARDAN_CLUSTER_01_Urban': [('00627299', 'North')],
    'MIRPUR_CLUSTER_01_Rural': [('00627299', 'North')],
    'MIRPUR_CLUSTER_01_Urban': [('00627299', 'North')],
    'MOTORWAY_CLUSTER_01_Rural': [('00627299', 'North')],
    'MURREE_CLUSTER_01_Rural': [('00627299', 'North')],
    'MURREE_CLUSTER_01_Urban': [('00627299', 'North')],
    'MUZAFFARABAD_CLUSTER_01_Rural': [('00627299', 'North')],
    'MUZAFFARABAD_CLUSTER_01_Urban': [('00627299', 'North')],
    'NOWSHEHRA_CLUSTER_01_Rural': [('00627299', 'North')],
    'NOWSHEHRA_CLUSTER_01_Urban': [('00627299', 'North')],
    'PESHAWAR_CLUSTER_01_Rural': [('00627299', 'North')],
    'PESHAWAR_CLUSTER_01_Urban': [('00627299', 'North')],
    'PESHAWAR_CLUSTER_02_Rural': [('00627299', 'North')],
    'PESHAWAR_CLUSTER_02_Urban': [('00627299', 'North')],
    'PESHAWAR_CLUSTER_03_Rural': [('00627299', 'North')],
    'PESHAWAR_CLUSTER_03_Urban': [('00627299', 'North')],
    'PESHAWAR_CLUSTER_04_Rural': [('00627299', 'North')],
    'PESHAWAR_CLUSTER_04_Urban': [('00627299', 'North')],
    'RAWALPINDI_CLUSTER_01_Rural': [('00627299', 'North')],
    'RAWALPINDI_CLUSTER_01_Urban': [('00627299', 'North')],
    'RAWALPINDI_CLUSTER_02_Rural': [('00627299', 'North')],
    'RAWALPINDI_CLUSTER_02_Urban': [('00627299', 'North')],
    'RAWALPINDI_CLUSTER_03_Urban': [('00627299', 'North')],
    'RAWALPINDI_CLUSTER_04_Rural': [('00627299', 'North')],
    'RAWALPINDI_CLUSTER_04_Urban': [('00627299', 'North')],
    'RAWALPINDI_CLUSTER_05_Rural': [('00627299', 'North')],
    'RAWALPINDI_CLUSTER_05_Urban': [('00627299', 'North')],
    'SWABI_CLUSTER_01_Rural': [('00627299', 'North')],
    'SWABI_CLUSTER_01_Urban': [('00627299', 'North')],
    'SWAT_CLUSTER_01_Rural': [('00627299', 'North')],
    'SWAT_CLUSTER_01_Urban': [('00627299', 'North')],
    'SWAT_CLUSTER_02_Rural': [('00627299', 'North')],
    'TALAGANG_CLUSTER_01_Rural': [('00627299', 'North')],
    'TAXILA_CLUSTER_01_Rural': [('00627299', 'North')],
    'TAXILA_CLUSTER_01_Urban': [('00627299', 'North')],
    # South Clusters
    'CHAMAN_CLUSTER_21_Rural': [('WX70154', 'South')],
    'CHAMAN_CLUSTER_21_Urban': [('WX70154', 'South')],
    'DADU_CLUSTER_15_Rural': [('WX70154', 'South')],
    'DADU_CLUSTER_15_Urban': [('WX70154', 'South')],
    'GAWADAR_CLUSTER_20_Rural': [('WX70154', 'South')],
    'GAWADAR_CLUSTER_20_Urban': [('WX70154', 'South')],
    'GHOTKI_CLUSTER_09_Rural': [('WX70154', 'South')],
    'GHOTKI_CLUSTER_09_Urban': [('WX70154', 'South')],
    'HYDERABAD_CLUSTER_01_RURAL': [('WX70154', 'South')],
    'HYDERABAD_CLUSTER_01_URBAN': [('WX70154', 'South')],
    'HYDERABAD_CLUSTER_02_RURAL': [('WX70154', 'South')],
    'HYDERABAD_CLUSTER_02_URBAN': [('WX70154', 'South')],
    'HYDERABAD_CLUSTER_03_RURAL': [('WX70154', 'South')],
    'HYDERABAD_CLUSTER_03_URBAN': [('WX70154', 'South')],
    'HYDERABAD_CLUSTER_04_RURAL': [('WX70154', 'South')],
    'HYDERABAD_CLUSTER_04_URBAN': [('WX70154', 'South')],
    'JACOBABAD_CLUSTER_12_Rural': [('WX70154', 'South')],
    'JACOBABAD_CLUSTER_12_Urban': [('WX70154', 'South')],
    'KARACHI_CLUSTER_01_RURAL': [('WX70154', 'South')],
    'KARACHI_CLUSTER_01_URBAN': [('WX70154', 'South')],         
    'KARACHI_CLUSTER_02_RURAL': [('WX70154', 'South')],
    'KARACHI_CLUSTER_02_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_03_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_04_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_05_RURAL': [('WX70154', 'South')],
    'KARACHI_CLUSTER_05_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_06_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_07_RURAL': [('WX70154', 'South')],
    'KARACHI_CLUSTER_07_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_08_URBAN': [('WX70154', 'South')],   
    'KARACHI_CLUSTER_09_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_10_RURAL': [('WX70154', 'South')],
    'KARACHI_CLUSTER_10_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_11_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_12_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_13_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_14_RURAL': [('WX70154', 'South')],
    'KARACHI_CLUSTER_14_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_15_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_16_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_17_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_18_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_19_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_20_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_21_RURAL': [('WX70154', 'South')],
    'KARACHI_CLUSTER_21_URBAN': [('WX70154', 'South')],
    'KARACHI_CLUSTER_22_RURAL': [('WX70154', 'South')],
    'KARACHI_CLUSTER_22_URBAN': [('WX70154', 'South')],
    'KHUZDAR_CLUSTER_17_Rural': [('WX70154', 'South')],
    'KHUZDAR_CLUSTER_17_Urban': [('WX70154', 'South')],
    'LARKANA_CLUSTER_13_Rural': [('WX70154', 'South')],
    'LARKANA_CLUSTER_13_Urban': [('WX70154', 'South')],
    'MIRPURKHAS_CLUSTER_05_Rural': [('WX70154', 'South')],
    'MIRPURKHAS_CLUSTER_05_Urban': [('WX70154', 'South')],
    'MITHI_CLUSTER_04_Rural': [('WX70154', 'South')],
    'MITHI_CLUSTER_04_Urban': [('WX70154', 'South')],
    'MORO_CLUSTER_02_Rural': [('WX70154', 'South')],
    'MORO_CLUSTER_02_Urban': [('WX70154', 'South')],
    'NAWABSHAH_CLUSTER_01_Rural': [('WX70154', 'South')],
    'NAWABSHAH_CLUSTER_01_Urban': [('WX70154', 'South')],
    'QUETTAEAST_CLUSTER_19_Urban': [('WX70154', 'South')],
    'QUETTAWEST_CLUSTER_18_Rural': [('WX70154', 'South')],
    'QUETTAWEST_CLUSTER_18_Urban': [('WX70154', 'South')],
    'SAKRAND_CLUSTER_06_Rural': [('WX70154', 'South')],
    'SAKRAND_CLUSTER_06_Urban': [('WX70154', 'South')],
    'SHAHDADKOT_CLUSTER_14_Rural': [('WX70154', 'South')],
    'SHAHDADKOT_CLUSTER_14_Urban': [('WX70154', 'South')],
    'SIBBI_CLUSTER_16_Rural': [('WX70154', 'South')],
    'SIBBI_CLUSTER_16_Urban': [('WX70154', 'South')],
    'SUI_CLUSTER_11_Rural': [('WX70154', 'South')],
    'SUI_CLUSTER_11_Urban': [('WX70154', 'South')],
    'SUKKUR_CLUSTER_10_Rural': [('WX70154', 'South')],
    'SUKKUR_CLUSTER_10_Urban': [('WX70154', 'South')],
    'THATTA_CLUSTER_07_Rural': [('WX70154', 'South')],
    'THATTA_CLUSTER_07_Urban': [('WX70154', 'South')],
    'UMERKOT_CLUSTER_03_Rural': [('WX70154', 'South')],
    'UMERKOT_CLUSTER_03_Urban': [('WX70154', 'South')],
     # Center-1 Clusters
    'GUJRANWALA_CLUSTER_01_Rural': [('WX410590', 'Center-1')],
    'GUJRANWALA_CLUSTER_01_Urban': [('WX410590', 'Center-1')],
    'GUJRANWALA_CLUSTER_02_Rural': [('WX1091941', 'Center-1')],
    'GUJRANWALA_CLUSTER_02_Urban': [('WX1091941', 'Center-1')],
    'GUJRANWALA_CLUSTER_03_Rural': [('WX410590', 'Center-1')],
    'GUJRANWALA_CLUSTER_03_Urban': [('WX410590', 'Center-1')],
    'GUJRANWALA_CLUSTER_04_Rural': [('WX410590', 'Center-1')],
    'GUJRANWALA_CLUSTER_04_Urban': [('WX410590', 'Center-1')],
    'GUJRANWALA_CLUSTER_05_Rural': [('WX410590', 'Center-1')],
    'GUJRANWALA_CLUSTER_05_Urban': [('WX410590', 'Center-1')],
    'GUJRANWALA_CLUSTER_06_Rural': [('WX832498', 'Center-1')], 
    'KASUR_CLUSTER_01_Rural': [('WX325496', 'Center-1')],
    'KASUR_CLUSTER_02_Rural': [('WX325496', 'Center-1')],
    'KASUR_CLUSTER_03_Rural': [('WX325496', 'Center-1')],
    'KASUR_CLUSTER_03_Urban': [('WX325496', 'Center-1')],
    'LAHORE_CLUSTER_01_Rural': [('WX410590', 'Center-1')],
    'LAHORE_CLUSTER_01_Urban': [('WX410590', 'Center-1')],
    'LAHORE_CLUSTER_02_Rural': [('WX410591', 'Center-1')],
    'LAHORE_CLUSTER_02_Urban': [('WX410591', 'Center-1')],
    'LAHORE_CLUSTER_03_Rural': [('WX410591', 'Center-1')],
    'LAHORE_CLUSTER_03_Urban': [('WX410591', 'Center-1')],
    'LAHORE_CLUSTER_04_Urban': [('WX410591', 'Center-1')],
    'LAHORE_CLUSTER_05_Rural': [('WX410591', 'Center-1')],
    'LAHORE_CLUSTER_05_Urban': [('WX410591', 'Center-1')],
    'LAHORE_CLUSTER_06_Rural': [('WX310192', 'Center-1')],
    'LAHORE_CLUSTER_06_Urban': [('WX310192', 'Center-1')],
    'LAHORE_CLUSTER_07_Rural': [('WX1295428', 'Center-1')],
    'LAHORE_CLUSTER_07_Urban': [('WX1295428', 'Center-1')],
    'LAHORE_CLUSTER_08_Rural': [('WX1295428', 'Center-1')],
    'LAHORE_CLUSTER_08_Urban': [('WX1295428', 'Center-1')],
    'LAHORE_CLUSTER_09_Rural': [('WX410591', 'Center-1')],
    'LAHORE_CLUSTER_09_Urban': [('WX410591', 'Center-1')],
    'LAHORE_CLUSTER_10_Urban': [('WX310192', 'Center-1')],
    'LAHORE_CLUSTER_11_Rural': [('WX310192', 'Center-1')],
    'LAHORE_CLUSTER_11_Urban': [('WX310192', 'Center-1')],
    'LAHORE_CLUSTER_12_Urban': [('WX1295428', 'Center-1')],
    'LAHORE_CLUSTER_13_Urban': [('WX1295428', 'Center-1')],
    'LAHORE_CLUSTER_14_Urban': [('WX310192', 'Center-1')],  
    'SIALKOT_CLUSTER_01_Rural': [('WX288666', 'Center-1')],
    'SIALKOT_CLUSTER_01_Urban': [('WX288666', 'Center-1')],
    'SIALKOT_CLUSTER_02_Rural': [('WX288666', 'Center-1')],
    'SIALKOT_CLUSTER_02_Urban': [('WX288666', 'Center-1')],
    'SIALKOT_CLUSTER_03_Rural': [('WX288666', 'Center-1')],
    'SIALKOT_CLUSTER_03_Urban': [('WX288666', 'Center-1')],
    'SIALKOT_CLUSTER_04_Rural': [('WX410590', 'Center-1')],
    'SIALKOT_CLUSTER_05_Rural': [('WX288666', 'Center-1')],
    'SIALKOT_CLUSTER_05_Urban': [('WX288666', 'Center-1')],
    'SIALKOT_CLUSTER_06_Rural': [('WX832498', 'Center-1')],
    'SIALKOT_CLUSTER_06_Urban': [('WX832498', 'Center-1')],
    'SIALKOT_CLUSTER_07_Rural': [('WX832498', 'Center-1')],
    'SIALKOT_CLUSTER_07_Urban': [('WX832498', 'Center-1')], 
    # Center-2 Clusters
    'DG_KHAN_CLUSTER_01_Rural': [('WX176265', 'Center-2')],
    'DG_KHAN_CLUSTER_02_Rural': [('WX176265', 'Center-2')],
    'DG_KHAN_CLUSTER_02_Urban': [('WX176265', 'Center-2')],
    'DI_KHAN_CLUSTER_01_Rural': [('WX176265', 'Center-2')],
    'DI_KHAN_CLUSTER_01_Urban': [('WX176265', 'Center-2')],
    'DI_KHAN_CLUSTER_02_Rural': [('WX176265', 'Center-2')],
    'DI_KHAN_CLUSTER_02_Urban': [('WX176265', 'Center-2')],
    'DI_KHAN_CLUSTER_03_Rural': [('WX529634', 'Center-2')],   
    'FAISALABAD_CLUSTER_01_Rural': [('WX344281', 'Center-2')],
    'FAISALABAD_CLUSTER_02_Rural': [('WX344281', 'Center-2')],
    'FAISALABAD_CLUSTER_03_Rural': [('WX344281', 'Center-2')],
    'FAISALABAD_CLUSTER_04_Rural': [('WX356127', 'Center-2')],
    'FAISALABAD_CLUSTER_04_Urban': [('WX356127', 'Center-2')],
    'FAISALABAD_CLUSTER_05_Rural': [('WX356127', 'Center-2')],
    'FAISALABAD_CLUSTER_05_Urban': [('WX356127', 'Center-2')],
    'FAISALABAD_CLUSTER_06_Rural': [('WX356127', 'Center-2')],
    'FAISALABAD_CLUSTER_06_Urban': [('WX356127', 'Center-2')], 
    'JHUNG_CLUSTER_01_Rural': [('WX356127', 'Center-2')],
    'JHUNG_CLUSTER_01_Urban': [('WX356127', 'Center-2')],
    'JHUNG_CLUSTER_02_Rural': [('WX356127', 'Center-2')],
    'JHUNG_CLUSTER_02_Urban': [('WX356127', 'Center-2')],
    'JHUNG_CLUSTER_03_Rural': [('WX529634', 'Center-2')],
    'JHUNG_CLUSTER_03_Urban': [('WX529634', 'Center-2')],
    'JHUNG_CLUSTER_04_Rural': [('WX529634', 'Center-2')],
    'JHUNG_CLUSTER_04_Urban': [('WX529634', 'Center-2')],
    'JHUNG_CLUSTER_05_Rural': [('WX529634', 'Center-2')],
    'JHUNG_CLUSTER_05_Urban': [('WX529634', 'Center-2')], 
    'SAHIWAL_CLUSTER_01_Rural': [('WX344281', 'Center-2')],
    'SAHIWAL_CLUSTER_01_Urban': [('WX344281', 'Center-2')],
    'SAHIWAL_CLUSTER_02_Rural': [('WX344281', 'Center-2')],
    'SAHIWAL_CLUSTER_02_Urban': [('WX344281', 'Center-2')],
    # Center-3 Clusters
    'AHMEDPUREAST_CLUSTER_01_Rural': [('WX480533', 'Center-3')],
    'AHMEDPUREAST_CLUSTER_01_Urban': [('WX480533', 'Center-3')],
    'ALIPUR_CLUSTER_01_Rural': [('WX434107', 'Center-3')],
    'ALIPUR_CLUSTER_01_Urban': [('WX434107', 'Center-3')],
    'BAHAWALPUR_CLUSTER_01_Rural': [('WX480533', 'Center-3')],
    'BAHAWALPUR_CLUSTER_01_Urban': [('WX480533', 'Center-3')],
    'BAHAWALPUR_CLUSTER_02_Rural': [('WX480533', 'Center-3')],  
    'DG_KHAN_CLUSTER_03_Rural': [('WX434107', 'Center-3')],
    'DG_KHAN_CLUSTER_03_Urban': [('WX434107', 'Center-3')],
    'DG_KHAN_CLUSTER_04_Rural': [('WX1286760', 'Center-3')],
    'DG_KHAN_CLUSTER_04_Urban': [('WX1286760', 'Center-3')], 
    'JAMPUR_CLUSTER_01_Rural': [('WX1286760', 'Center-3')],
    'JAMPUR_CLUSTER_01_Urban': [('WX1286760', 'Center-3')],
    'KHANPUR_CLUSTER_01_Rural': [('WX480533', 'Center-3')],
    'KHANPUR_CLUSTER_01_Urban': [('WX480533', 'Center-3')], 
    'MULTAN_CLUSTER_01_Rural': [('WX434107', 'Center-3')],
    'MULTAN_CLUSTER_01_Urban': [('WX434107', 'Center-3')],
    'MULTAN_CLUSTER_02_Rural': [('WX434107', 'Center-3')],
    'MULTAN_CLUSTER_02_Urban' : [('WX434107', 'Center-3')],
    'MULTAN_CLUSTER_03_Rural': [('WX434107', 'Center-3')],
    'MULTAN_CLUSTER_03_Urban' : [('WX434107', 'Center-3')],
    'RAHIMYARKHAN_CLUSTER_01_Rural' : [('WX480533', 'Center-3')],
    'RAHIMYARKHAN_CLUSTER_01_Urban': [('WX480533', 'Center-3')],
    'RAJANPUR_CLUSTER_01_Rural': [('WX1286760', 'Center-3')],
    'RAJANPUR_CLUSTER_01_Urban': [('WX1286760', 'Center-3')],
    'SADIQABAD_CLUSTER_01_Rural': [('WX1286760', 'Center-3')],
    'SAHIWAL_CLUSTER_03_Rural': [('WX278447', 'Center-3')],
    'SAHIWAL_CLUSTER_03_Urban': [('WX278447', 'Center-3')],
    'SAHIWAL_CLUSTER_04_Rural': [('WX278447', 'Center-3')],
    'SAHIWAL_CLUSTER_04_Urban': [('WX278447', 'Center-3')],
    'SAHIWAL_CLUSTER_05_Rural': [('WX278447', 'Center-3')],
    'SAHIWAL_CLUSTER_05_Urban': [('WX278447', 'Center-3')],
    'SAHIWAL_CLUSTER_06_Rural': [('WX278447', 'Center-3')],
    'SAHIWAL_CLUSTER_06_Urban': [('WX278447', 'Center-3')],
    'SAHIWAL_CLUSTER_07_Rural': [('WX278447', 'Center-3')],
    'SAHIWAL_CLUSTER_07_Urban': [('WX278447', 'Center-3')]}


# convert the dict to Pandas DataFrame
df6 = pd.DataFrame([(cluster, values[0][0], values[0][1]) for cluster, values in data.items()],
                  columns=['NE Information;Cluster', 'NE Information;Engineer', 'NE Information;Region'])
df6.head()
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
      <th>NE Information;Cluster</th>
      <th>NE Information;Engineer</th>
      <th>NE Information;Region</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>ABBOTTABAD_CLUSTER_01_Rural</td>
      <td>00627299</td>
      <td>North</td>
    </tr>
    <tr>
      <th>1</th>
      <td>ABBOTTABAD_CLUSTER_01_Urban</td>
      <td>00627299</td>
      <td>North</td>
    </tr>
    <tr>
      <th>2</th>
      <td>AJK_CLUSTER_01_Rural</td>
      <td>00627299</td>
      <td>North</td>
    </tr>
    <tr>
      <th>3</th>
      <td>AJK_CLUSTER_01_Urban</td>
      <td>00627299</td>
      <td>North</td>
    </tr>
    <tr>
      <th>4</th>
      <td>BANNU_CLUSTER_01_Rural</td>
      <td>00627299</td>
      <td>North</td>
    </tr>
  </tbody>
</table>
</div>




```python
final_df.head()
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
      <th>NE Information;Cluster</th>
      <th>NE Information;BSCName</th>
      <th>NE Information;CI</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>31516</td>
    </tr>
    <tr>
      <th>1</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>33498</td>
    </tr>
    <tr>
      <th>2</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>21516</td>
    </tr>
    <tr>
      <th>3</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>23498</td>
    </tr>
    <tr>
      <th>4</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>11516</td>
    </tr>
  </tbody>
</table>
</div>




```python
df7 = pd.merge(final_df,df6,how='left',on=['NE Information;Cluster'])
```


```python
df7
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
      <th>NE Information;Cluster</th>
      <th>NE Information;BSCName</th>
      <th>NE Information;CI</th>
      <th>NE Information;Engineer</th>
      <th>NE Information;Region</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>31516</td>
      <td>WX70154</td>
      <td>South</td>
    </tr>
    <tr>
      <th>1</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>33498</td>
      <td>WX70154</td>
      <td>South</td>
    </tr>
    <tr>
      <th>2</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>21516</td>
      <td>WX70154</td>
      <td>South</td>
    </tr>
    <tr>
      <th>3</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>23498</td>
      <td>WX70154</td>
      <td>South</td>
    </tr>
    <tr>
      <th>4</th>
      <td>CHAMAN_CLUSTER_21_Rural</td>
      <td>HQTABSC04</td>
      <td>11516</td>
      <td>WX70154</td>
      <td>South</td>
    </tr>
    <tr>
      <th>...</th>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <th>46490</th>
      <td>SIALKOT_CLUSTER_07_Urban</td>
      <td>HLHRBSC02</td>
      <td>11733</td>
      <td>WX832498</td>
      <td>Center-1</td>
    </tr>
    <tr>
      <th>46491</th>
      <td>SIALKOT_CLUSTER_07_Urban</td>
      <td>HLHRBSC02</td>
      <td>17471</td>
      <td>WX832498</td>
      <td>Center-1</td>
    </tr>
    <tr>
      <th>46492</th>
      <td>SIALKOT_CLUSTER_07_Urban</td>
      <td>HLHRBSC02</td>
      <td>31283</td>
      <td>WX832498</td>
      <td>Center-1</td>
    </tr>
    <tr>
      <th>46493</th>
      <td>SIALKOT_CLUSTER_07_Urban</td>
      <td>HLHRBSC02</td>
      <td>11283</td>
      <td>WX832498</td>
      <td>Center-1</td>
    </tr>
    <tr>
      <th>46494</th>
      <td>SIALKOT_CLUSTER_07_Urban</td>
      <td>HLHRBSC02</td>
      <td>21283</td>
      <td>WX832498</td>
      <td>Center-1</td>
    </tr>
  </tbody>
</table>
<p>46495 rows × 5 columns</p>
</div>



## PRS Cluster Output


```python
# Set Output Path
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(folder_path)
# Export
df7.to_csv('03_PRS.csv',index=False)
```

## Re-Set Variables


```python
# re-set all the variable from the RAM
%reset -f
```

## Delte Un-wanted Folders


```python
import os
import shutil
working_directory = 'D:/Advance_Data_Sets/PRS/GSM'
os.chdir(working_directory)
# Python Script for Clearing Subfolders in a Specific path
for subfolder in os.listdir(working_directory):
    subfolder_path = os.path.join(working_directory, subfolder)
    if os.path.isdir(subfolder_path):
        shutil.rmtree(subfolder_path)
```


```python
# re-set all the variable from the RAM
%reset -f
```


```python

```
