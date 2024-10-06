#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# Daily Conformance

### Import required Libraries


```python
import os
import zipfile
import numpy as np
import pandas as pd
from glob import glob
from collections import ChainMap
```

### Set Working Path


```python
folder_path = 'D:/Advance_Data_Sets/SLA/Cluster_BH_KPIs'
os.chdir(folder_path)
```

### Import Cluster BH KPIs


```python
df = sorted(glob('*.zip'))
```


```python
# concat all the Cluster DA Files
df0=pd.concat((pd.read_csv(file,skiprows=range(6),encoding='unicode_escape',\
        skipfooter=1,engine='python',\
        usecols=['Date','Time','GCell Group',\
                'CSSR_Non Blocking','DCR','Out_HSR',\
                'RxQual Index DL(%)','RxQual Index UL(%)',\
                'GOS-SDCCH(%)','CallSetup TCH GOS(%)','Mobility TCH GOS(%)'],\
        parse_dates=["Date"],na_values=['NIL','/0']) for file in df)).\
        sort_values('Date').reset_index(drop=True).\
        rename(columns={'CSSR_Non Blocking': 'CSSR',\
                       'DCR': 'DCR',\
                       'Out_HSR': 'HSR',\
                       'RxQual Index DL(%)': 'DL RQI',\
                       'RxQual Index UL(%)': 'UL RQI',\
                       'GOS-SDCCH(%)': 'SDCCH GoS',\
                       'CallSetup TCH GOS(%)': 'TCH GoS',\
                       'Mobility TCH GOS(%)': 'MoB GoS'}).fillna('N/A')
```

### Filter only Clusters


```python
# Filter only cluster , remove city, region, sub region
df1=df0[df0['GCell Group']\
        .str.contains('|'.join(['_Rural','_Urban','RURAL','_URBAN']))]\
        .copy().reset_index(drop=True)
```

### Sub Region Defination


```python
df2 = ChainMap(dict.fromkeys(['GUJRANWALA_CLUSTER_01_Rural',
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
                            'KASUR_CLUSTER_01_Rural',
                            'KASUR_CLUSTER_02_Rural',
                            'KASUR_CLUSTER_03_Rural',
                            'KASUR_CLUSTER_03_Urban',
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
                            'SIALKOT_CLUSTER_07_Urban'], 'Center-1'), 
                            dict.fromkeys(['DG_KHAN_CLUSTER_01_Rural',
                            'DG_KHAN_CLUSTER_02_Rural',
                            'DG_KHAN_CLUSTER_02_Urban',
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
                            'SAHIWAL_CLUSTER_01_Rural',
                            'SAHIWAL_CLUSTER_01_Urban',
                            'SAHIWAL_CLUSTER_02_Rural',
                            'SAHIWAL_CLUSTER_02_Urban'], 'Center-2'), 
                            dict.fromkeys(['JAMPUR_CLUSTER_01_Urban',
                           'RAJANPUR_CLUSTER_01_Rural',
                           'RAJANPUR_CLUSTER_01_Urban',
                           'JAMPUR_CLUSTER_01_Rural',
                           'DG_KHAN_CLUSTER_03_Rural',
                           'DG_KHAN_CLUSTER_03_Urban',
                           'DG_KHAN_CLUSTER_04_Rural',
                           'DG_KHAN_CLUSTER_04_Urban',
                           'SAHIWAL_CLUSTER_03_Rural',
                           'SAHIWAL_CLUSTER_03_Urban',
                           'KHANPUR_CLUSTER_01_Rural',
                           'KHANPUR_CLUSTER_01_Urban',
                           'RAHIMYARKHAN_CLUSTER_01_Rural',
                           'RAHIMYARKHAN_CLUSTER_01_Urban',
                           'AHMEDPUREAST_CLUSTER_01_Rural',
                           'AHMEDPUREAST_CLUSTER_01_Urban',
                           'ALIPUR_CLUSTER_01_Rural',
                           'ALIPUR_CLUSTER_01_Urban',
                           'BAHAWALPUR_CLUSTER_01_Rural',
                           'BAHAWALPUR_CLUSTER_01_Urban',
                           'BAHAWALPUR_CLUSTER_02_Rural',
                           'SAHIWAL_CLUSTER_04_Rural',
                           'SAHIWAL_CLUSTER_04_Urban',
                           'MULTAN_CLUSTER_01_Rural',
                           'MULTAN_CLUSTER_01_Urban',
                           'MULTAN_CLUSTER_02_Rural',
                           'MULTAN_CLUSTER_02_Urban',
                           'MULTAN_CLUSTER_03_Rural',
                           'MULTAN_CLUSTER_03_Urban',
                           'RYK DESERT_Cluster_Rural',
                           'SADIQABAD_CLUSTER_01_Rural',
                           'SAHIWAL_CLUSTER_05_Rural',
                           'SAHIWAL_CLUSTER_05_Urban',
                           'SAHIWAL_CLUSTER_06_Rural',
                           'SAHIWAL_CLUSTER_06_Urban',
                           'SAHIWAL_CLUSTER_07_Rural',
                           'SAHIWAL_CLUSTER_07_Urban'], 'Center-3'),
                            dict.fromkeys(['ABBOTTABAD_CLUSTER_01_Rural',
                                           'ABBOTTABAD_CLUSTER_01_Urban',
                                           'AJK_CLUSTER_01_Rural',
                                           'AJK_CLUSTER_01_Urban',
                                           'BANNU_CLUSTER_01_Rural',
                                           'CHAKWAL_CLUSTER_01_Rural',
                                           'CHAKWAL_CLUSTER_01_Urban',
                                           'FANA_CLUSTER_01_Rural',
                                           'GTROAD_CLUSTER_01_Rural',
                                           'GTROAD_CLUSTER_01_Urban',
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
                                           'JHELUM_CLUSTER_01_Rural',
                                           'JHELUM_CLUSTER_01_Urban',
                                           'KOHAT_CLUSTER_01_Rural',
                                           'KOHAT_CLUSTER_01_Urban',
                                           'KOHAT_CLUSTER_02_Rural',
                                           'MANSEHRA_CLUSTER_01_Rural',
                                           'MANSEHRA_CLUSTER_01_Urban',
                                           'MARDAN_CLUSTER_01_Rural',
                                           'MARDAN_CLUSTER_01_Urban',
                                           'MIRPUR_CLUSTER_01_Rural',
                                           'MIRPUR_CLUSTER_01_Urban',
                                           'MOTORWAY_CLUSTER_01_Rural',
                                           'MURREE_CLUSTER_01_Rural',
                                           'MURREE_CLUSTER_01_Urban',
                                           'MUZAFFARABAD_CLUSTER_01_Rural',
                                           'MUZAFFARABAD_CLUSTER_01_Urban',
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
                                           'RAWALPINDI_CLUSTER_01_Rural',
                                           'RAWALPINDI_CLUSTER_01_Urban',
                                           'RAWALPINDI_CLUSTER_02_Rural',
                                           'RAWALPINDI_CLUSTER_02_Urban',
                                           'RAWALPINDI_CLUSTER_03_Urban',
                                           'RAWALPINDI_CLUSTER_04_Rural',
                                           'RAWALPINDI_CLUSTER_04_Urban',
                                           'RAWALPINDI_CLUSTER_05_Rural',
                                           'RAWALPINDI_CLUSTER_05_Urban',
                                           'SWABI_CLUSTER_01_Rural',
                                           'SWABI_CLUSTER_01_Urban',
                                           'SWAT_CLUSTER_01_Rural',
                                           'SWAT_CLUSTER_01_Urban',
                                           'SWAT_CLUSTER_02_Rural',
                                           'TALAGANG_CLUSTER_01_Rural',
                                           'TAXILA_CLUSTER_01_Rural',
                                           'TAXILA_CLUSTER_01_Urban'], 'North'),
                            dict.fromkeys(['CHAMAN_CLUSTER_21_Rural',
                            'CHAMAN_CLUSTER_21_Urban',
                            'DADU_CLUSTER_15_Rural',
                            'DADU_CLUSTER_15_Urban',
                            'GAWADAR_CLUSTER_20_Rural',
                            'GAWADAR_CLUSTER_20_Urban',
                            'GHOTKI_CLUSTER_09_Rural',
                            'GHOTKI_CLUSTER_09_Urban',
                            'HYDERABAD_CLUSTER_01_RURAL',
                            'HYDERABAD_CLUSTER_01_URBAN',
                            'HYDERABAD_CLUSTER_02_RURAL',
                            'HYDERABAD_CLUSTER_02_URBAN',
                            'HYDERABAD_CLUSTER_03_RURAL',
                            'HYDERABAD_CLUSTER_03_URBAN',
                            'HYDERABAD_CLUSTER_04_RURAL',
                            'HYDERABAD_CLUSTER_04_URBAN',
                            'JACOBABAD_CLUSTER_12_Rural',
                            'JACOBABAD_CLUSTER_12_Urban',
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
                            'KHUZDAR_CLUSTER_17_Rural',
                            'KHUZDAR_CLUSTER_17_Urban',
                            'LARKANA_CLUSTER_13_Rural',
                            'LARKANA_CLUSTER_13_Urban',
                            'MIRPURKHAS_CLUSTER_05_Rural',
                            'MIRPURKHAS_CLUSTER_05_Urban',
                            'MITHI_CLUSTER_04_Rural',
                            'MITHI_CLUSTER_04_Urban',
                            'MORO_CLUSTER_02_Rural',
                            'MORO_CLUSTER_02_Urban',
                            'NAWABSHAH_CLUSTER_01_Rural',
                            'NAWABSHAH_CLUSTER_01_Urban',
                            'QUETTAEAST_CLUSTER_19_Urban',
                            'QUETTAWEST_CLUSTER_18_Rural',
                            'QUETTAWEST_CLUSTER_18_Urban',
                            'SAKRAND_CLUSTER_06_Rural',
                            'SAKRAND_CLUSTER_06_Urban',
                            'SHAHDADKOT_CLUSTER_14_Rural',
                            'SHAHDADKOT_CLUSTER_14_Urban',
                            'SIBBI_CLUSTER_16_Rural',
                            'SIBBI_CLUSTER_16_Urban',
                            'SUI_CLUSTER_11_Rural',
                            'SUI_CLUSTER_11_Urban',
                            'SUKKUR_CLUSTER_10_Rural',
                            'SUKKUR_CLUSTER_10_Urban',
                            'THATTA_CLUSTER_07_Rural',
                            'THATTA_CLUSTER_07_Urban',
                            'UMERKOT_CLUSTER_03_Rural',
                            'UMERKOT_CLUSTER_03_Urban'], 'South'))

df1.loc[:, 'Region'] = df1['GCell Group'].map(df2.get)
```

### Modification Cluster Name as per requirement


```python
# Cluster is Urban or Rural, get from GCell Group
df1['Cluster Type'] = df1['GCell Group'].str.strip().str[-5:]
```


```python
# Remove Urban and Rural from GCell Group
df1['Location']=df1['GCell Group'].map(lambda x: str(x)[:-6])
```

### Select Required Quater


```python
# Identify the Quater 
df1['Quater'] = pd.PeriodIndex(pd.to_datetime(df1.Date), freq='Q')
```


```python
# Select Required Quater
df3=df1[df1.Quater=='2024Q4'].copy()
```


```python
df3.loc[:, 'Cluster Type'] = df3['Cluster Type'].str.capitalize()
```

### Re-Shape the DataSet


```python
df4=df3.pivot(index=['Date','Region','Location'],\
        columns='Cluster Type',\
        values=['CSSR','DCR','HSR',\
               'SDCCH GoS','TCH GoS','MoB GoS',\
               'DL RQI','UL RQI']).\
         fillna('N/A').\
        sort_index(level=[0,1],\
        axis=1,ascending=[True,False])
```

### Required Sequenc


```python
#Rrequired sequence
df5=df4[[('CSSR', 'Urban'),('CSSR', 'Rural'),\
           ('DCR', 'Urban'),('DCR', 'Rural'),\
            ('HSR', 'Urban'),('HSR', 'Rural'),\
            ('SDCCH GoS', 'Urban'),('SDCCH GoS', 'Rural'),\
            ('TCH GoS', 'Urban'),('TCH GoS', 'Rural'),\
            ('MoB GoS', 'Urban'),('MoB GoS', 'Rural'),\
            ('DL RQI', 'Urban'),('DL RQI', 'Rural'),\
            ('UL RQI', 'Urban'),('UL RQI', 'Rural')]].reset_index()
```

### Export Output


```python
with pd.ExcelWriter('SLA_Target.xlsx',\
                    date_format = 'dd-mm-yyyy',\
                    datetime_format='dd-mm-yyyy') as writer:
    # GSM KPIs Values
    df5.to_excel(writer,engine='xlsxwriter',\
                 na_rep='N/A',sheet_name="Summary")
```


```python
# re-set all the variable from the RAM
%reset -f
```

### Formatting Excel File


```python
import os
folder_path = 'D:/Advance_Data_Sets/SLA/Cluster_BH_KPIs'
os.chdir(folder_path)
```


```python
import openpyxl
# Load the workbook
wb = openpyxl.load_workbook('SLA_Target.xlsx')
```

### Select the Required Sheet


```python
# Select the worksheet
ws = wb['Summary']
```

### Delete the Row


```python
ws.delete_rows(3)
```

### Set Tab Color


```python
# Set the tab color
ws.sheet_properties.tabColor = "A020F0" 
```

### Set Date Column Format


```python
from datetime import datetime
from openpyxl.styles import NamedStyle

# Assuming you have the worksheet (ws) and the date column index is 2 (column B)
date_column_index = 2

# Define a custom date style
date_style = NamedStyle(name='custom_date_style', number_format='DD/MM/YYYY')

for row in ws.iter_rows(min_col=date_column_index, max_col=date_column_index):
    for cell in row:
        if isinstance(cell.value, datetime):
            cell.style = date_style
```

### Set Float Number Format


```python
for row in ws.iter_rows(min_col=5, max_col=ws.max_column):
    for cell in row:
        if isinstance(cell.value,float):
            cell.number_format = '0.00'
```

### Merge Row


```python
# Merge cells in rows B1:B2
ws.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)
# Merge cells in rows C1:C2
ws.merge_cells(start_row=1, start_column=3, end_row=2, end_column=3)
# Merge cells in rows D1:D2
ws.merge_cells(start_row=1, start_column=4, end_row=2, end_column=4)
```

### Apply Border


```python
# Define the border style
border = openpyxl.styles.borders.Border(
    left=openpyxl.styles.borders.Side(style='thin'),
    right=openpyxl.styles.borders.Side(style='thin'),
    top=openpyxl.styles.borders.Side(style='thin'),
    bottom=openpyxl.styles.borders.Side(style='thin'))

# Loop through each cell in the worksheet and set the border style
for row in ws.iter_rows():
    for cell in row:
        cell.border = border
```


```python
for row in ws.iter_rows():
    for cell in row:
        cell.font = openpyxl.styles.Font(name='Calibri Light', size=11)
```

### Header Format


```python
# set background and font color
for row in range(1,3):
    for cell in ws[row]:
        cell.fill = openpyxl.styles.PatternFill(start_color="C00000",\
                                                end_color="C00000", \
                                                fill_type = "solid")
        font = openpyxl.styles.Font(color="FFFFFF",\
                                    bold=True,\
                                    size=11,\
                                    name='Calibri Light')
        cell.font = font
```

### Apply Alignment


```python
# Loop through each cell in the worksheet and set the alignment
for row in ws.iter_rows():
    for cell in row:
        cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
```

### Set Column width


```python
from openpyxl.utils import get_column_letter

# Set the width of column B to 10
ws.column_dimensions[get_column_letter(2)].width = 12

# Set the width of column C to 8
ws.column_dimensions[get_column_letter(3)].width = 10

# Set the width of column D to 26
ws.column_dimensions[get_column_letter(4)].width = 26


# Set the width of all other columns to 8
for col in range(6,ws.max_column+1):
    ws.column_dimensions[get_column_letter(col)].width = 8
```

### Hide Column


```python
# hide the 1st column
ws.column_dimensions.group('A', hidden=True)
```

### Conditional Formatting


```python
from openpyxl.formatting.rule import CellIsRule
# Define a variable for the fill color for empty cells
empty_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')

# Iterate through all the rows and columns
for row in ws.rows:
    for cell in row:
        if cell.value in (None, 'N/A', ' ',):
            cell.fill = empty_color
```

### 1. CSSR Urban


```python
# Define range_string
last_row = ws.max_row
```


```python
range_string_E = "E3:E" + str(last_row)   # CSSR-Urban

# Define the conditional formatting rules for E column (CSSR URBAN)
ws.conditional_formatting.add(range_string_E, \
                               CellIsRule(operator='lessThanOrEqual', \
                               formula=[99.3], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws.conditional_formatting.add(range_string_E, \
                               CellIsRule(operator='between', \
                               formula=[99.3,99.5], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_E, \
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[99.5],\
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 2. CSSR Rural


```python
range_string_F = "F3:F" + str(last_row)   # CSSR-Rural
# Define the conditional formatting rules for F column (CSSR_RURAL)
ws.conditional_formatting.add(range_string_F, \
                               CellIsRule(operator='lessThanOrEqual', \
                               formula=[98.7], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws.conditional_formatting.add(range_string_F, \
                               CellIsRule(operator='between', \
                                          formula=[98.7,99.0], \
                                          fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_F, \
                               CellIsRule(operator='greaterThanOrEqual', \
                               formula=[99.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 3. DCR Urban


```python
range_string_G = "G3:G" + str(last_row)   # DCR-Urban

# Define the conditional formatting rules for F column (DCR-Urban)
ws.conditional_formatting.add(range_string_G, \
                            CellIsRule(operator='lessThanOrEqual', \
                            formula=[0.6], \
                            fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws.conditional_formatting.add(range_string_G,\
                               CellIsRule(operator='between', \
                               formula=[0.6,0.65], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_G, \
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[0.65], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 4. DCR Rural


```python
range_string_H = "H3:H" + str(last_row)   # DCR-Rural
# Define the conditional formatting rules for G column (DCR-Rural)
ws.conditional_formatting.add(range_string_H,\
                               CellIsRule(operator='lessThanOrEqual', \
                            formula=[1.0], \
                            fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws.conditional_formatting.add(range_string_H, \
                            CellIsRule(operator='between', \
                            formula=[1.0,1.05], \
                            fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_H, \
                               CellIsRule(operator='greaterThanOrEqual', \
                               formula=[1.05], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 5. HSR Urban 


```python
range_string_I = "I3:I" + str(last_row)   # HSR-Urban

# Define the conditional formatting rules for H column (HSR-Urban)
ws.conditional_formatting.add(range_string_I, \
                               CellIsRule(operator='lessThanOrEqual', \
                                formula=[97], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws.conditional_formatting.add(range_string_I, \
                            CellIsRule(operator='between', \
                            formula=[97,97.5], \
                            fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_I, \
                                CellIsRule(operator='greaterThanOrEqual', \
                                formula=[97.5], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 6. HSR Rural 


```python
range_string_J = "J3:J" + str(last_row)   # HSR-Rural
# Define the conditional formatting rules for I column (HSR-Rural)
ws.conditional_formatting.add(range_string_J, \
                               CellIsRule(operator='lessThanOrEqual', \
                                formula=[95.6], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws.conditional_formatting.add(range_string_J, \
                               CellIsRule(operator='between', formula=[95.6,96], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_J, \
                               CellIsRule(operator='greaterThanOrEqual', \
                               formula=[96.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 7. SDCCH GoS-Urban


```python
range_string_K = "K3:K" + str(last_row)   # SDCCH GoS-Urban

# Define the conditional formatting rules for J column (SDCCH GoS-Urban)
ws.conditional_formatting.add(range_string_K, \
                               CellIsRule(operator='lessThanOrEqual', \
                               formula=[0.12], \
                               fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws.conditional_formatting.add(range_string_K, \
                               CellIsRule(operator='between', \
                               formula=[0.1,0.12], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_K, \
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[0.12], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 8. SDCCH GoS-Rural


```python
range_string_L = "L3:L" + str(last_row)   # SDCCH GoS-Rural
# Define the conditional formatting rules for K column (SDCCH GoS-Rural)
ws.conditional_formatting.add(range_string_L, \
                               CellIsRule(operator='lessThanOrEqual', \
                               formula=[0.1], \
                               fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))


ws.conditional_formatting.add(range_string_L, \
                               CellIsRule(operator='between', \
                                formula=[0.1,0.12], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))


ws.conditional_formatting.add(range_string_L, \
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[0.12], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 9.TCH GoS-Urban


```python
range_string_M = "M3:M" + str(last_row)   # TCH GoS-Urban

# Define the conditional formatting rules for L column (TCH GoS-Urban)
ws.conditional_formatting.add(range_string_M, \
                               CellIsRule(operator='lessThanOrEqual', \
                                formula=[2.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws.conditional_formatting.add(range_string_M, \
                               CellIsRule(operator='between', \
                                formula=[2.,2.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_M, \
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[2.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 10.TCH GoS-Rural


```python
range_string_N = "N3:N" + str(last_row)   # TCH GoS-Rural
# Define the conditional formatting rules for M column (TCH GoS-Rural)
ws.conditional_formatting.add(range_string_N, \
                               CellIsRule(operator='lessThanOrEqual', \
                                          formula=[2.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws.conditional_formatting.add(range_string_N, \
                               CellIsRule(operator='between', \
                                formula=[2,2.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_N, \
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[2.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 11. MoB GoS-Urban


```python
range_string_O = "O3:O" + str(last_row)   # MoB GoS-Urban

# Define the conditional formatting rules for N column (MoB GoS-Urban)
ws.conditional_formatting.add(range_string_O, \
                               CellIsRule(operator='lessThanOrEqual', \
                                formula=[4.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws.conditional_formatting.add(range_string_O, \
                               CellIsRule(operator='between', \
                                formula=[4.0,4.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_O, \
                               CellIsRule(operator='greaterThanOrEqual',\
                                formula=[4.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 12. MoB GoS-Rural


```python
range_string_P = "P3:P" + str(last_row)   # MoB GoS-Rural
# Define the conditional formatting rules for O column (MoB GoS-Rural)
ws.conditional_formatting.add(range_string_P, \
                               CellIsRule(operator='lessThanOrEqual', \
                                formula=[4.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws.conditional_formatting.add(range_string_P, \
                               CellIsRule(operator='between', \
                                formula=[4.0,4.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_P, \
                               CellIsRule(operator='greaterThanOrEqual',\
                                formula=[4.2], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 13. DL RQI-Urban


```python
range_string_Q = "Q3:Q" + str(last_row)   # DL RQI-Urban

# Define the conditional formatting rules for P column (DL RQI-Urban)
ws.conditional_formatting.add(range_string_Q, \
                               CellIsRule(operator='lessThanOrEqual',\
                               formula=[98.0], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws.conditional_formatting.add(range_string_Q, \
                               CellIsRule(operator='between', \
                                formula=[98.0,98.4],\
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_Q,\
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[98.4],\
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 14. DL RQI-Rural


```python
range_string_R = "R3:R" + str(last_row)   # DL RQI-Rural
# Define the conditional formatting rules for Q column (DL RQI-Rural)
ws.conditional_formatting.add(range_string_R, \
                               CellIsRule(operator='lessThanOrEqual', \
                               formula=[96.5], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws.conditional_formatting.add(range_string_R, \
                               CellIsRule(operator='between', \
                               formula=[96.5,97], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_R,\
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[97.0], \
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 15. UL RQI-Urban


```python
range_string_S = "S3:S" + str(last_row)   # UL RQI-Urban

# Define the conditional formatting rules for R column (UL RQI-Urban)
ws.conditional_formatting.add(range_string_S, \
                               CellIsRule(operator='lessThanOrEqual', \
                                formula=[98.0], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws.conditional_formatting.add(range_string_S, \
                               CellIsRule(operator='between', \
                               formula=[98.0,98.2], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_S, \
                               CellIsRule(operator='greaterThanOrEqual', \
                               formula=[98.2], \
                               fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 16. UL RQI-Rural


```python
range_string_T = "T3:T" + str(last_row)   # UL RQI-Rural
# Define the conditional formatting rules for S column (UL RQI-Rural)
ws.conditional_formatting.add(range_string_T,\
                               CellIsRule(operator='lessThanOrEqual',\
                               formula=[97.2], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws.conditional_formatting.add(range_string_T, \
                               CellIsRule(operator='between', \
                               formula=[97.2,97.7], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws.conditional_formatting.add(range_string_T, \
                               CellIsRule(operator='greaterThanOrEqual', \
                               formula=[97.7], \
                               fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### Set Zoom Size


```python
# Apply the zoomScale
ws.sheet_view.zoomScale = 80
```

### Set Filter on the Header


```python
# Get the first row
first_row = ws[1]
# Apply the filter on the first row
ws.auto_filter.ref = f"A2:{get_column_letter(len(first_row))}2"
```

## Insert a New Sheet (as First Sheet)


```python
ws511 =wb.create_sheet("Title Page", 0)
```

## Merge Specific Row and Columns


```python
ws511.merge_cells(start_row=12, start_column=5, end_row=18, end_column=17)
```

## Fill the Merge Cells


```python
# Fill the text 'CE Utilization Report' in the merged cell
ws511.cell(row=12, column=5).value = 'Cluster Busy Hour KPIs'
```

## Formatting Tital Page Report


```python
# Access the first row starting from row 3
first_row1 = list(ws511.rows)[11]
# Iterate through the cells in the first row starting from column E
for cell in first_row1[4:]:
    cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
    cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", end_color="CF0A2C", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=45,name='Calibri Light')
    cell.font = font
```

## Inset Image


```python
from openpyxl.drawing.image import Image

# inset the Huawei logo
img = Image('D:/Advance_Data_Sets/KPIs_Analysis/Huawei.jpg')
img.width = 7 * 15
img.height = 7 * 15
ws511.add_image(img,'E3')

# inset the PTCL logo
img1 = Image('D:/Advance_Data_Sets/KPIs_Analysis/PTCL.png')
ws511.add_image(img1,'M3')
```

## Hide the gridlines


```python
ws511.sheet_view.showGridLines = False
```

## Hide the headings


```python
ws511.sheet_view.showRowColHeaders = False
```

## Hyper Link For Title Page


```python
# loop through all sheets in the workbook and insert the hyperlink to each sheet
row = 22
for ws in wb:
    if ws.title != "Title Page":
        hyperlink_cell = ws511.cell(row=row, column=5)
        hyperlink_cell.value = ws.title
        hyperlink_cell.hyperlink = "#'{}'!A1".format(ws.title)
        hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
        hyperlink_cell.border = border
        hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        # set the height of the cell
        ws511.row_dimensions[row].height = 15
        # set the colum width of column 5
        ws511.column_dimensions[get_column_letter(5)].width = 30
        row += 1
```

## Hyper Link For Sub Pages


```python
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.styles import borders
# Loop through all sheets in the workbook and insert the hyperlink to each sheet
for i, ws in enumerate(wb.worksheets):
    # Check if the sheet is not the Title Page
    if ws.title != "Title Page":
        # Add hyperlink to cell in the last column+2 of the sheet
        hyperlink_cell = ws.cell(row=2, column=ws.max_column+2)
        hyperlink_cell.value = "Back to Table of Contents"
        hyperlink_cell.hyperlink = "#'{}'!E{}".format("Title Page", 22)
        hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
        hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        
        # Set border on the hyperlink column
        for row in ws.iter_rows(min_row=2, max_row=4, min_col=hyperlink_cell.column, max_col=hyperlink_cell.column):
            for cell in row:
                cell.border = border
        # Set width of the hyperlink column
        col_letter = openpyxl.utils.get_column_letter(hyperlink_cell.column)
        ws.column_dimensions[col_letter].width = 25

        # Add hyperlink to cell in the last column+2 of the sheet for next sheet
        if i < len(wb.worksheets)-1:
            next_hyperlink_cell = ws.cell(row=3, column=ws.max_column)
            next_hyperlink_cell.value = "Next Sheet"
            next_hyperlink_cell.hyperlink = "#'{}'!A1".format(wb.worksheets[i+1].title)
            next_hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
            next_hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        
        # Add hyperlink to cell in the last column+2 of the sheet for previous sheet
        if i > 0:
            prev_hyperlink_cell = ws.cell(row=4, column=ws.max_column)
            prev_hyperlink_cell.value = "Previous Sheet"
            prev_hyperlink_cell.hyperlink = "#'{}'!A1".format(wb.worksheets[i-1].title)
            prev_hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
            prev_hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
```

## Clear the Cell Value


```python
# Select the sheet by index
wss = wb.worksheets[1]
# Get the cell at row 4 and the last column
cell = wss.cell(row=4, column=wss.max_column)
# Clear the contents of the cell
cell.value = None
# Remove the border of the cell
cell.border = None
```

### Final Output


```python
# Save the changes
wb.save('SLA_Target.xlsx')
```

### Re-Set Variables


```python
# re-set all the variable from the RAM
%reset -f
```

## Move Final Output to Ouput Folder


```python
import shutil

# set the file path and folder paths
file_path = "D:/Advance_Data_Sets/SLA/Cluster_BH_KPIs/SLA_Target.xlsx"
destination_folder = "D:/Advance_Data_Sets/Output_Folder"

# use the shutil.move() function to move the file
shutil.move(file_path, destination_folder)
```
