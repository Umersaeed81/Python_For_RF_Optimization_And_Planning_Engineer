#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

## Import Libraries


```python
# Python libraries
import os
import pandas as pd
```

## Set Input File Path


```python
# Set Working Folder Path
path = 'C:/Users/UWX161178/Downloads/BCP'
os.chdir(path)
```

## Import Input File


```python
df1 = pd.read_excel('Freq_Plan_18072024.xlsx',sheet_name=0)
```

## Convert to List 


```python
df1.loc[:, 'TCH of BCCH Module']  = df1['TCH of BCCH Module'].str.split(';').apply(list)
```

## Re-Shape Data Set


```python
df2=df1.explode('TCH of BCCH Module').reset_index(level=-1, drop=True)
```

## Export Output


```python
df2.to_csv('output.csv',index=False)
```
