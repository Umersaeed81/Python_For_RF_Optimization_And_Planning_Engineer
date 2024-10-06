#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# LTE RF Export (Cell File) Processing

## Import Required Libraries


```python
import os
import glob
import numpy as np
import pandas as pd
import warnings
warnings.simplefilter("ignore")
```

## User Define Funcation to import CELL File


```python
def read_excel_file(file_path):
    try:
        # Try to read 'CELL' sheet
        df = pd.read_excel(
            file_path,
            sheet_name='CELL',
            header=1,
            usecols=['*eNodeB Name', '*Cell Name', 'Cell active state', 'Downlink EARFCN', '*Local Cell ID'],
            dtype={'*Local Cell ID': str},
            engine='openpyxl'
        )
    except ValueError:
        # If 'CELL' sheet is not found, read 'LTE Cell' sheet
        df = pd.read_excel(
            file_path,
            sheet_name='LTE Cell',
            header=1,
            usecols=['*eNodeB Name', '*Cell Name', 'Cell active state', 'Downlink EARFCN', '*LocalCellID'],
            dtype={'*LocalCellID': str},
            engine='openpyxl'
        )
        # Rename the column to match the first dataframe's column name
        df = df.rename(columns={'*LocalCellID': '*Local Cell ID'})
    
    # Replace multiple spaces with single space in 'eNodeB Name' column
    df['*eNodeB Name'] = df['*eNodeB Name'].str.replace(r'\s+', ' ', regex=True)
    # Replace multiple spaces with single space in 'Cell Name' column
    df['*Cell Name'] = df['*Cell Name'].str.replace(r'\s+', ' ', regex=True)
    return df
```

## Import and Concatenate all Files


```python
df0= sorted(glob.glob('D:/Advance_Data_Sets/RF_Export/LTE/Cell/*.xlsx'))
df1 = pd.concat([read_excel_file(file) for file in df0]).drop_duplicates().reset_index(drop=True)
```

## Replace multiple spaces with single space in 'eNodeB Name' & 'CELLNAME' column


```python
# Replace multiple spaces with single space in 'Cell Name' column
df1['Cell Name'] = df1['*Cell Name'].str.replace(r'\s+', ' ', regex=True)
# Replace multiple spaces with single space in 'NodeB Name' column
df1['eNodeB Name'] = df1['*eNodeB Name'].str.replace(r'\s+', ' ', regex=True)
```

## Generate LTE Site/Cell IDs


```python
# Generate LTE Site ID
df1['LTE_Site_ID'] = df1['*eNodeB Name'].str.split('_').str[0]
# Generate LTE Cell ID
df1['LTE_Cell_ID'] = df1['*Cell Name'].str.split('_').str[0]
```

## Generate Key


```python
# Extract data after the last hyphen and add it as a new column
df1['LTE_Key'] = df1['LTE_Site_ID']+";"+(df1['LTE_Cell_ID'].apply(lambda x: x.split('-')[-1])).astype(str)
```

## Band Identification


```python
df1['Band'] = np.where(
            (df1['Downlink EARFCN'].eq(1276)),
            'PTML-L-1800', 
            np.where(
                    (df1['Downlink EARFCN'].eq(3628)), 
                    'PTML-L-900', 
                np.where(
                    (df1['Downlink EARFCN'].eq(175)), 
                    'PTML-L-2100', 
                     'Other-Band')))
```

## Identify Cell Names that are duplicated


```python
#Identify Cell Names that are duplicated
df2 = df1[df1.duplicated(subset='*Cell Name', keep=False)].reset_index(drop=True)
```

## Filter ACTIVATED Cells From Duplicate Data Frame 


```python
df3 = df2[df2['Cell active state'].eq('CELL_ACTIVE')]
```

## Filter out the duplicate Cell Names entirely


```python
# Filter out the duplicate Cell Names entirely
df4 = df1[~df1['*Cell Name'].isin(df3['*Cell Name'])].reset_index(drop=True)
```

## Append df3 and df4


```python
# Append df2 and df3
df5 = pd.concat([df3, df4], ignore_index=True)
```

## Replace Values


```python
# Replace values based on conditions using np.where
df5['LTE_Key'] = np.where(
    df5['Band'] == 'PTML-L-2100', 
    df5['LTE_Key'].replace({ ';81': ';1', 
                             ';82': ';2',
                             ';83': ';3',
                             ';84': ';4',
                             ';85': ';5',
                             ';86': ';6',
                             ';87': ';7',
                             ';88': ';8',
                             ';89': ';9'}, regex=True),
    np.where(
        df5['Band'] == 'PTML-L-900',
        df5['LTE_Key'].replace({';21': ';1', 
                                ';22': ';2',
                                ';23': ';3',
                                ';24': ';4', 
                                ';41': ';1',
                                ';42': ';2',
                                ';43': ';3',
                                ';44': ';4'
                               }, regex=True),
        df5['LTE_Key']
    ))
```

## Filter Requied Band


```python
df6 = df5[df5['Band'].ne('Other-Band')]
```

## Export LTE RF Export Output


```python
folder_path = 'D:/Advance_Data_Sets/GUL/GUL_Output'
os.chdir(folder_path)
df6.to_csv('04_LTE_RF_Export.csv',index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
