# Finding Closest Points Using Cartesian Product and Haversine Formula

## Import Required Libraries


```python
import os
import itertools
import numpy as np
import pandas as pd
from math import radians, cos, sin, sqrt, atan2
```

## Set Input File Path


```python
folder_path = 'D:/Advance_Data_Sets/Distance'
os.chdir(folder_path)
```

## Import Input Data Set


```python
table_a = pd.read_csv('Ookla_File.csv')
```


```python
table_a.head()
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
      <th>test_id</th>
      <th>client_latitude</th>
      <th>client_longitude</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>10347511113</td>
      <td>24.999</td>
      <td>44.989</td>
    </tr>
    <tr>
      <th>1</th>
      <td>10346353590</td>
      <td>17.320</td>
      <td>43.306</td>
    </tr>
    <tr>
      <th>2</th>
      <td>10347096683</td>
      <td>27.859</td>
      <td>42.428</td>
    </tr>
    <tr>
      <th>3</th>
      <td>10346204575</td>
      <td>24.466</td>
      <td>39.617</td>
    </tr>
    <tr>
      <th>4</th>
      <td>10347607777</td>
      <td>27.859</td>
      <td>42.428</td>
    </tr>
  </tbody>
</table>
</div>




```python
table_b = pd.read_excel('EP.xlsx')
```


```python
table_b.head()
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
      <th>Site ID</th>
      <th>Lon</th>
      <th>Lat</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>RIY0030</td>
      <td>46.627175</td>
      <td>24.901619</td>
    </tr>
    <tr>
      <th>1</th>
      <td>RIY0337</td>
      <td>46.787950</td>
      <td>24.964622</td>
    </tr>
    <tr>
      <th>2</th>
      <td>RIY0371</td>
      <td>46.754080</td>
      <td>24.722260</td>
    </tr>
    <tr>
      <th>3</th>
      <td>RIY0388</td>
      <td>46.837000</td>
      <td>24.784870</td>
    </tr>
    <tr>
      <th>4</th>
      <td>RIY0390</td>
      <td>46.808970</td>
      <td>24.905530</td>
    </tr>
  </tbody>
</table>
</div>



## Cartesian product


```python
# Perform Cartesian product
product = np.array(list(itertools.product(table_a.values, table_b.values)))
```

## DataFrame


```python
# Unpack the data directly
test_id, lat_a, lon_a = product[:, 0, 0], product[:, 0, 1], product[:, 0, 2]
site_id, lon_b, lat_b = product[:, 1, 0], product[:, 1, 1], product[:, 1, 2]

# Create the DataFrame with the unpacked arrays
df_result = pd.DataFrame({
    'test_id': test_id,
    'Lat_l': lat_a,
    'Lon_l': lon_a,
    'Site ID': site_id,
    'Lon_s': lon_b,
    'Lat_s': lat_b
})
```

## Pre-Processing


```python
# Ensure columns are numeric (convert to float if necessary)
df_result['Lat_l'] = pd.to_numeric(df_result['Lat_l'], errors='coerce')
df_result['Lon_l'] = pd.to_numeric(df_result['Lon_l'], errors='coerce')
df_result['Lat_s'] = pd.to_numeric(df_result['Lat_s'], errors='coerce')
df_result['Lon_s'] = pd.to_numeric(df_result['Lon_s'], errors='coerce')
```

## User Define Funcation


```python
# Haversine function with vectorized operations
def haversine_vectorized(lat1, lon1, lat2, lon2):
    # Radius of the Earth in kilometers
    R = 6371.0  
    
    # Convert latitude and longitude from degrees to radians
    lat1 = np.radians(lat1)
    lon1 = np.radians(lon1)
    lat2 = np.radians(lat2)
    lon2 = np.radians(lon2)
    
    # Differences in coordinates
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    
    # Haversine formula
    a = np.sin(dlat / 2)**2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon / 2)**2
    c = 2 * np.arcsin(np.sqrt(a))
    
    # Distance in kilometers
    return R * c
```

## Calculate Distance


```python
# Use the vectorized Haversine function
df_result['Distance_km'] = haversine_vectorized(df_result['Lat_l'], df_result['Lon_l'], \
                                                df_result['Lat_s'], df_result['Lon_s'])
```

## Calculate Min. Distance


```python
df_min = df_result.loc[df_result.groupby(['test_id'])\
            ['Distance_km'].idxmin()].reset_index(drop=True)
```

## Export Oput


```python
df_min.to_csv('Min_Distance_Using_haversine.csv',index=False)
```
