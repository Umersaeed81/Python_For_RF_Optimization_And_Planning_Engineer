#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# 2G CGID Logs Analysis

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

## Set CGID Log File Path


```python
#set the Path (Path must be same format)
path = 'D:/Advance_Data_Sets/TXN_DataSets/GSM_CGID_log_Analysis'
os.chdir(path)
```

## Get the file name from the Folder and Sub-Folder


```python
matchers = ['.log_']
nicFiles = [file for file in glob(path + '/**/*', recursive=True) \
            if any(xs in file for xs in matchers) and os.path.isfile(file)]
```

## Import CGID Log Files


```python
df0 = (pd.read_csv(f,header=None,usecols=[0,6,12],parse_dates=["Time"],\
            names=['Time','Type','Cell Index'])\
            .assign(File=f.split('.')[0]) for f in nicFiles)

# concatenate files
df1 = pd.concat(df0, ignore_index=True)
```

## Data Pre-Processing


```python
#Preprocessing to get the BSC NE
df1['NE FDN'] = df1['File'].str.split(' ').\
                    str[1].str.split('\\').str[0].\
                    str.split('_').str[1]
```


```python
# Data Pre-Processing
df1['Cell Index']=df1['Cell Index'].str.replace("CellId:","")
df1['Time']=df1['Time'].str.replace("Time:","")
df1['Time']=df1['Time'].astype('datetime64[ns]') 
df1['Time'] = df1['Time'].dt.date  # Extract date component
df1['Time'] = pd.to_datetime(df1['Time'])  # Convert to datetime object
df1['Cell Index']=df1['Cell Index'].astype(str)
df1 = df1.sort_values(by='Time', ascending=True).reset_index(drop=True)
```

## Set Time Column as Index


```python
df1.set_index('Time', inplace=True)
```

## Get the current date


```python
# Get the current date
today = pd.to_datetime('today').normalize()
```

## Filter Required Data


```python
# Filter the data to include only the last 7 days (excluding today)
df2 = df1[df1.index < today].last('7D').reset_index()
```

## Import BSC NE Info File


```python
df3 = pd.read_csv('TimeCostReport.csv',\
                  usecols=['NE FDN','NE Name']).\
                  drop_duplicates().dropna().\
                  reset_index(drop=True)
```

## Merge CGID logs and BSC NE Files


```python
df4 = pd.merge(df2,df3[['NE FDN','NE Name']],on=['NE FDN'])
```

## Set RF Export (Fequency Export) Path


```python
folder_path= 'D:/Advance_Data_Sets/RF_Export/GSM/00_Frequency_Export'
os.chdir(folder_path)
```

## Unzip File


```python
for file in os.listdir(folder_path):   # get the list of files
    if zipfile.is_zipfile(file): # if it is a zipfile, extract it
        with zipfile.ZipFile(file) as zip_ref: # treat the file as a zip
            folder_name = os.path.splitext(file)[0]  # get the name of the folder (remove '.zip' extension)
            os.makedirs(folder_name, exist_ok=True)  # create the folder if it doesn't exist
            zip_ref.extractall(folder_name)  # extract the contents into the folder
```

## Import GCELL File


```python
# List Required File
df5 = glob('D:/Advance_Data_Sets/RF_Export/GSM/00_Frequency_Export/**/GCELL.txt', recursive=True)
# import and concat gcell
df6=pd.concat((pd.read_csv(file, encoding='unicode_escape',\
                            header=1,\
                            usecols=['BSC Name','BTS Name','Cell Index',\
                             'Cell Name','Cell CI'],\
                            dtype = {"Cell Index" : "str"})) for file in df5).\
                            rename(columns = {"BSC Name":"NE Name"})
```

## Merge CGID log with GCell File (From RF Export)


```python
df7 = pd.merge(df4,df6,on=['NE Name','Cell Index'])
```

## Filter only Mute Type


```python
df8 = df7[df7['Type'].eq('Type:Mute')].reset_index(drop=True)
```

## Re-Shape(pivot_table) CGID Log Analysis


```python
df9 = df8.pivot_table\
      (index=['NE Name','Cell Index','BTS Name','Cell Name'],\
      columns="Time",values='File',aggfunc='count',\
      margins=True,margins_name='Total Mutating').iloc[:-1, :].\
      fillna(0).reset_index()
```

## Summary w.r.t BSC


```python
df10 = df8.pivot_table\
      (index=['NE Name'],\
      columns="Time",values='File',aggfunc='count',\
      margins=True,margins_name='Total Mutating').iloc[:-1, :].\
      fillna(0).reset_index()
```

## Export


```python
folder_path= 'D:/Advance_Data_Sets/BSS_EI_Output/CDIG_Processing'
os.chdir(folder_path)
```


```python
with pd.ExcelWriter('00_CDIG_Analysis.xlsx',date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    # Cell Levek
    df9.to_excel(writer,sheet_name="CDIG_Cell_Level",index=False)
    # BSC Level
    df10.to_excel(writer,sheet_name="CDIG_BSC_Level",index=False)
```


```python
# re-set all the variable from the RAM
%reset -f
```
