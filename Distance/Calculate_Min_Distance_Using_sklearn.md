#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>


# Efficient Nearest Neighbor Search Using BallTree and Haversine Distance

The objective of this code is to find the closest geographical point from one dataset to each point in another dataset using an efficient algorithm. Here’s a simple breakdown of what the code does:

1. **Import Necessary Libraries:** It imports libraries for data handling, mathematical operations, and spatial analysis.

2. **Set the Working Directory:** The code changes the working directory to a specified folder that contains the datasets.

3. **Load Data:** It reads data from a CSV file (`Ookla_File.csv`) and an Excel file (`EP.xlsx`), which include geographical coordinates (latitude and longitude).

4. **Convert Coordinates to Radians:** It converts the latitude and longitude coordinates from degrees to radians for accurate distance calculations.

5. **Prepare BallTree Structure:** The code prepares a BallTree structure using the latitude and longitude data from the second dataset (`table_b`). BallTree is a data structure that helps efficiently find nearest neighbors in multi-dimensional spaces.

6. **Find Closest Points:** It queries the BallTree to find the closest point in `table_b` for each point in `table_a`. It retrieves the distances and the indices of the closest points.

7. **Convert Distance Units:** The distances obtained are in radians, so the code converts them to kilometers (considering the Earth's radius).

8. **Store Results:** It adds two new columns to `table_a`: one for the closest site ID from `table_b` and another for the corresponding distance in kilometers.

9. **Prepare Final Output:** Finally, it creates a new DataFrame containing relevant information, including test ID, client coordinates, closest site ID, and closest distance.

Overall, this code helps identify the nearest geographical locations from a second dataset for each location in the first dataset, facilitating geographical analysis and decision-making.

## Import Required Libraries


```python
import os
import numpy as np
import pandas as pd
from sklearn.neighbors import BallTree
```

## Set Input File Path


```python
folder_path = 'D:/Advance_Data_Sets/Distance'
os.chdir(folder_path)
```

## Import Input Data Set


```python
table_a = pd.read_csv('Ookla_File.csv')
table_a['Lat_rad'] = np.radians(table_a['client_latitude'])
table_a['Lon_rad'] = np.radians(table_a['client_longitude'])
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
      <th>Lat_rad</th>
      <th>Lon_rad</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>10347511113</td>
      <td>24.999</td>
      <td>44.989</td>
      <td>0.436315</td>
      <td>0.785206</td>
    </tr>
    <tr>
      <th>1</th>
      <td>10346353590</td>
      <td>17.320</td>
      <td>43.306</td>
      <td>0.302291</td>
      <td>0.755832</td>
    </tr>
    <tr>
      <th>2</th>
      <td>10347096683</td>
      <td>27.859</td>
      <td>42.428</td>
      <td>0.486231</td>
      <td>0.740508</td>
    </tr>
    <tr>
      <th>3</th>
      <td>10346204575</td>
      <td>24.466</td>
      <td>39.617</td>
      <td>0.427012</td>
      <td>0.691447</td>
    </tr>
    <tr>
      <th>4</th>
      <td>10347607777</td>
      <td>27.859</td>
      <td>42.428</td>
      <td>0.486231</td>
      <td>0.740508</td>
    </tr>
  </tbody>
</table>
</div>




```python
table_b = pd.read_excel('EP.xlsx')
table_b['Lat_rad'] = np.radians(table_b['Lat'])
table_b['Lon_rad'] = np.radians(table_b['Lon'])
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
      <th>Lat_rad</th>
      <th>Lon_rad</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>RIY0030</td>
      <td>46.627175</td>
      <td>24.901619</td>
      <td>0.434615</td>
      <td>0.813798</td>
    </tr>
    <tr>
      <th>1</th>
      <td>RIY0337</td>
      <td>46.787950</td>
      <td>24.964622</td>
      <td>0.435715</td>
      <td>0.816604</td>
    </tr>
    <tr>
      <th>2</th>
      <td>RIY0371</td>
      <td>46.754080</td>
      <td>24.722260</td>
      <td>0.431485</td>
      <td>0.816013</td>
    </tr>
    <tr>
      <th>3</th>
      <td>RIY0388</td>
      <td>46.837000</td>
      <td>24.784870</td>
      <td>0.432578</td>
      <td>0.817460</td>
    </tr>
    <tr>
      <th>4</th>
      <td>RIY0390</td>
      <td>46.808970</td>
      <td>24.905530</td>
      <td>0.434684</td>
      <td>0.816971</td>
    </tr>
  </tbody>
</table>
</div>



## Nearest Site Identification and Distance Calculation Based on Geospatial Coordinates


```python
# Prepare the BallTree with Table B's lat/lon in radians
tree = BallTree(np.c_[table_b['Lat_rad'], table_b['Lon_rad']], metric='haversine')

# Query the closest point in Table A for each point in Table B
distances, indices = tree.query(np.c_[table_a['Lat_rad'], table_a['Lon_rad']], k=1)

# Convert the distance from radians to kilometers (radius of Earth = 6371 km)
table_a['Closest Site ID'] = table_b['Site ID'].iloc[indices.flatten()].values
table_a['Closest Distance (km)'] = distances.flatten() * 6371

# Display the result with the correct column names
min_distance= table_a[['test_id','client_latitude', 'client_longitude', 'Closest Site ID', 'Closest Distance (km)']]
```

## Export Output


```python
min_distance.to_csv('min_distance_using_nbr.csv',index=False)
```
