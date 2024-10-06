#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

## TCH Availability Rate(%) (BTS-DA Level)

## Import Required Libraries


```python
import os
#import glob
import zipfile
import pandas as pd
from glob import glob
import warnings
warnings.simplefilter("ignore")
```

## Delete Files From the Folder, If Folder Not Found Created the Folder


```python
# Folder path
folder_path = r"D:/Advance_Data_Sets/BSS_EI_Output/CDIG_Processing"

# Check if the folder exists
if os.path.exists(folder_path):
    # Iterate over the files in the folder and delete them
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted file: {file_path}")
        except Exception as e:
            print(f"Error deleting {file_path}: {e}")
else:
    # If the folder doesn't exist, create it
    print(f"Folder {folder_path} does not exist. Creating...")
    os.makedirs(folder_path)
    print(f"Folder {folder_path} created.")
```

## Set input File Path


```python
# set file path
working_directory = 'D:/Advance_Data_Sets/TXN_DataSets/GSM_BTS_DA'
os.chdir(working_directory)
```

## Import anc Concat BTS DA KPIs


```python
# list and filter required files 
df0 = sorted(glob('*.zip'))

# concat all the BTS DA Files
df1=pd.concat((pd.read_csv(file,skiprows=range(6),\
            skipfooter=1,engine='python',\
            usecols=['Date','GBSC','Site Name','TCH Availability Rate(%)'],
            parse_dates=["Date"],na_values=['NIL','/0']) for file in df0))\
            .sort_values('Date').set_index(['Date']).\
            last('7D').reset_index()
```

## Re-Share Data Set (Pivot_table)


```python
# Re-Shape Data Set
df2 = df1.pivot_table(index=['GBSC','Site Name'],\
                          columns='Date',values='TCH Availability Rate(%)').reset_index()
```

## Conditional Count


```python
# Conditional Count
df2["Number_of_Days_Availability<100"] = df2.where(df2.iloc[:,2:].\
                                  lt(100)).count(axis=1)
```

## Export


```python
folder_path= 'D:/Advance_Data_Sets/BSS_EI_Output/CDIG_Processing'
os.chdir(folder_path)
```


```python
with pd.ExcelWriter('01_CDIG_Analysis.xlsx',date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    # GSM Outage BTS Level
    df2.to_excel(writer,sheet_name="GSM_BTS_Ava",engine='openpyxl',index=False)
```


```python

```
