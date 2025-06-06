#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# Finding Closest Points Using Cartesian Product and Haversine Formula

The objective of the code is to calculate the shortest distance between two sets of geographical points using their latitude and longitude coordinates. Here's a breakdown of what the code does:

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Distance/example_1_b.png?raw=true)

1. **Import Necessary Libraries:** It imports libraries for handling data and performing mathematical calculations.

2. **Set the Working Directory:** The code changes the working directory to a specified folder containing the datasets.

3. **Load Data:** It reads data from a CSV file (`Ookla_File.csv`) and an Excel file (`EP.xlsx`), which contain geographical information.

4. **Create Cartesian Product:** It generates all possible combinations (Cartesian product) of the points from the two datasets, essentially pairing each point in the first dataset with each point in the second dataset.

5. **Unpack Data:** It extracts the relevant columns (test ID, latitude, and longitude) from the combined data for easier processing.

6. **Ensure Numeric Data:** It converts the latitude and longitude columns to numeric types, ensuring that any non-numeric entries are handled appropriately.

7. **Define Haversine Function:** It defines a function to calculate the great-circle distance (the shortest distance over the earth's surface) between two points given their latitude and longitude using the Haversine formula.

8. **Calculate Distances:** It applies the Haversine function to calculate the distance between each pair of points from the two datasets.

9. **Find Minimum Distances:** Finally, it identifies the closest point from the second dataset for each unique test ID in the first dataset and stores these results in a new DataFrame.

Overall, this code helps in analyzing distances between geographical locations from two different sources, allowing for further insights or decisions based on proximity.

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
