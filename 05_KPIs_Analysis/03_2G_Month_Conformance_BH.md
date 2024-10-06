#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# Monthly SLA tt

## Import required Libraries


```python
import os
import zipfile
import numpy as np
import pandas as pd
from glob import glob
from collections import ChainMap
```

## Set Working Path


```python
folder_path = 'D:/Advance_Data_Sets/SLA/Cluster_BH_KPIs'
os.chdir(folder_path)
```

## Import Cluster BH Counters


```python
df = sorted(glob('*.zip'))
# concat all the Cluster DA Files
df0=pd.concat((pd.read_csv(file,skiprows=range(6),encoding='unicode_escape',
        skipfooter=1,engine='python',
        usecols=['Date','Time','GCell Group',
        '_CallSetup TCH GOS(%)_D','_CallSetup TCH GOS(%)_N',
        '_GOS-SDCCH(%)_D','_GOS-SDCCH(%)_N',
        '_Mobility TCH GOS(%)_D','_Mobility TCH GOS(%)_N',
        '_DCR_D','_DCR_N',
        '_RxQual Index DL_1','_RxQual Index DL_2',
        '_RxQual Index UL_1','_RxQual Index UL_2',
        'N_OutHSR_D','N_OutHSR_N',
        'CSSR_Non Blocking_1_N','CSSR_Non Blocking_1_D',
        'CSSR_Non Blocking_2_N','CSSR_Non Blocking_2_D','GlobelTraffic','TCH Availability Rate(%)'],
        parse_dates=["Date"]) for file in df)).\
        sort_values('Date').reset_index(drop=True)
```

## Assign Month Number


```python
df0['Month']=df0['Date'].dt.month.astype(str).apply(lambda x: x.zfill(2))\
                                                            +"_"+ \
                                        df0['Date'].dt.year.astype(str)
```

## Filter Out un-required Cluster


```python
# Filter out un-required cluster
df1=df0[df0['GCell Group'].ne('RYK DESERT_Cluster_Rural')].copy().reset_index(drop=True)
```

## Sub Region Defination


```python
df2 = ChainMap(dict.fromkeys(['GUJRANWALA_CLUSTER_01_Rural',
                            'GUJRANWALA_CLUSTER_01_Urban',
                            'GUJRANWALA_CLUSTER_02_Rural','GUJRANWALA_CLUSTER_02_Urban',
                            'GUJRANWALA_CLUSTER_03_Rural','GUJRANWALA_CLUSTER_03_Urban',
                            'GUJRANWALA_CLUSTER_04_Rural','GUJRANWALA_CLUSTER_04_Urban',
                            'GUJRANWALA_CLUSTER_05_Rural','GUJRANWALA_CLUSTER_05_Urban',
                            'GUJRANWALA_CLUSTER_06_Rural','KASUR_CLUSTER_01_Rural',
                            'KASUR_CLUSTER_02_Rural','KASUR_CLUSTER_03_Rural',
                            'KASUR_CLUSTER_03_Urban','LAHORE_CLUSTER_01_Rural',
                            'LAHORE_CLUSTER_01_Urban',
                            'LAHORE_CLUSTER_02_Rural','LAHORE_CLUSTER_02_Urban',
                            'LAHORE_CLUSTER_03_Rural','LAHORE_CLUSTER_03_Urban',
                            'LAHORE_CLUSTER_04_Urban','LAHORE_CLUSTER_05_Rural',
                            'LAHORE_CLUSTER_05_Urban',
                            'LAHORE_CLUSTER_06_Rural','LAHORE_CLUSTER_06_Urban',
                            'LAHORE_CLUSTER_07_Rural','LAHORE_CLUSTER_07_Urban',
                            'LAHORE_CLUSTER_08_Rural','LAHORE_CLUSTER_08_Urban',
                            'LAHORE_CLUSTER_09_Rural','LAHORE_CLUSTER_09_Urban',
                            'LAHORE_CLUSTER_10_Urban','LAHORE_CLUSTER_11_Rural',
                            'LAHORE_CLUSTER_11_Urban','LAHORE_CLUSTER_12_Urban',
                            'LAHORE_CLUSTER_13_Urban','LAHORE_CLUSTER_14_Urban',
                            'SIALKOT_CLUSTER_01_Rural','SIALKOT_CLUSTER_01_Urban',
                            'SIALKOT_CLUSTER_02_Rural','SIALKOT_CLUSTER_02_Urban',
                            'SIALKOT_CLUSTER_03_Rural','SIALKOT_CLUSTER_03_Urban',
                            'SIALKOT_CLUSTER_04_Rural','SIALKOT_CLUSTER_05_Rural',
                            'SIALKOT_CLUSTER_05_Urban','SIALKOT_CLUSTER_06_Rural',
                            'SIALKOT_CLUSTER_06_Urban','SIALKOT_CLUSTER_07_Rural',
                            'SIALKOT_CLUSTER_07_Urban','LAHORE_City','SIALKOT_City',
                            'REGIONAL_CENTERAL','REGIONAL_CENTER01_215','REGIONAL_CENTERAL_From_July2024'], 'Center-1'), 
                            dict.fromkeys(['DG_KHAN_CLUSTER_01_Rural',
                            'DG_KHAN_CLUSTER_02_Rural','DG_KHAN_CLUSTER_02_Urban',
                            'DI_KHAN_CLUSTER_01_Rural','DI_KHAN_CLUSTER_01_Urban',
                            'DI_KHAN_CLUSTER_02_Rural',
                            'DI_KHAN_CLUSTER_02_Urban','DI_KHAN_CLUSTER_03_Rural',
                            'FAISALABAD_CLUSTER_01_Rural',
                            'FAISALABAD_CLUSTER_02_Rural','FAISALABAD_CLUSTER_03_Rural',
                            'FAISALABAD_CLUSTER_04_Rural',
                            'FAISALABAD_CLUSTER_04_Urban','FAISALABAD_CLUSTER_05_Rural',
                            'FAISALABAD_CLUSTER_05_Urban',
                            'FAISALABAD_CLUSTER_06_Rural','FAISALABAD_CLUSTER_06_Urban',
                            'JHUNG_CLUSTER_01_Rural',
                            'JHUNG_CLUSTER_01_Urban','JHUNG_CLUSTER_02_Rural',
                            'JHUNG_CLUSTER_02_Urban',
                            'JHUNG_CLUSTER_03_Rural','JHUNG_CLUSTER_03_Urban',
                           'JHUNG_CLUSTER_04_Rural',
                            'JHUNG_CLUSTER_04_Urban','JHUNG_CLUSTER_05_Rural',
                            'JHUNG_CLUSTER_05_Urban',
                            'SAHIWAL_CLUSTER_01_Rural','SAHIWAL_CLUSTER_01_Urban',
                            'SAHIWAL_CLUSTER_02_Rural','SAHIWAL_CLUSTER_02_Urban',
                            'FAISALABAD_City','REGIONAL_CENTER02_102'], 'Center-2'), 
                            dict.fromkeys(['JAMPUR_CLUSTER_01_Urban','RAJANPUR_CLUSTER_01_Rural',
                           'RAJANPUR_CLUSTER_01_Urban',
                            'JAMPUR_CLUSTER_01_Rural','DG_KHAN_CLUSTER_03_Rural',
                            'DG_KHAN_CLUSTER_03_Urban',
                            'DG_KHAN_CLUSTER_04_Rural','DG_KHAN_CLUSTER_04_Urban',
                            'SAHIWAL_CLUSTER_03_Rural',
                            'SAHIWAL_CLUSTER_03_Urban','KHANPUR_CLUSTER_01_Rural',
                           'KHANPUR_CLUSTER_01_Urban',
                            'RAHIMYARKHAN_CLUSTER_01_Rural','RAHIMYARKHAN_CLUSTER_01_Urban',
                            'AHMEDPUREAST_CLUSTER_01_Rural','AHMEDPUREAST_CLUSTER_01_Urban',
                            'ALIPUR_CLUSTER_01_Rural','ALIPUR_CLUSTER_01_Urban',
                            'BAHAWALPUR_CLUSTER_01_Rural','BAHAWALPUR_CLUSTER_01_Urban',
                            'BAHAWALPUR_CLUSTER_02_Rural','SAHIWAL_CLUSTER_04_Rural',
                            'SAHIWAL_CLUSTER_04_Urban','MULTAN_CLUSTER_01_Rural',
                            'MULTAN_CLUSTER_01_Urban','MULTAN_CLUSTER_02_Rural',
                            'MULTAN_CLUSTER_02_Urban',
                            'MULTAN_CLUSTER_03_Rural','MULTAN_CLUSTER_03_Urban',
                            'RYK DESERT_Cluster_Rural',
                            'SADIQABAD_CLUSTER_01_Rural','SAHIWAL_CLUSTER_05_Rural',
                           'SAHIWAL_CLUSTER_05_Urban',
                            'SAHIWAL_CLUSTER_06_Rural','SAHIWAL_CLUSTER_06_Urban',
                            'SAHIWAL_CLUSTER_07_Rural','SAHIWAL_CLUSTER_07_Urban',
                            'MULTAN_City','REGIONAL_CENTER03_102'], 'Center-3'),
                            dict.fromkeys(['ABBOTTABAD_CLUSTER_01_Rural','ABBOTTABAD_CLUSTER_01_Urban',
                            'AJK_CLUSTER_01_Rural','AJK_CLUSTER_01_Urban',
                            'BANNU_CLUSTER_01_Rural',
                            'CHAKWAL_CLUSTER_01_Rural','CHAKWAL_CLUSTER_01_Urban',
                            'FANA_CLUSTER_01_Rural',
                            'GTROAD_CLUSTER_01_Rural','GTROAD_CLUSTER_01_Urban',
                            'ISLAMABAD_CLUSTER_01_Rural','ISLAMABAD_CLUSTER_01_Urban',
                            'ISLAMABAD_CLUSTER_02_Rural','ISLAMABAD_CLUSTER_02_Urban',
                            'ISLAMABAD_CLUSTER_03_Rural','ISLAMABAD_CLUSTER_03_Urban',
                            'ISLAMABAD_CLUSTER_04_Rural','ISLAMABAD_CLUSTER_04_Urban',
                            'ISLAMABAD_CLUSTER_05_Rural','ISLAMABAD_CLUSTER_05_Urban',
                            'JHELUM_CLUSTER_01_Rural','JHELUM_CLUSTER_01_Urban',
                            'KOHAT_CLUSTER_01_Rural','KOHAT_CLUSTER_01_Urban','KOHAT_CLUSTER_02_Rural',
                            'MANSEHRA_CLUSTER_01_Rural',
                            'MANSEHRA_CLUSTER_01_Urban',
                            'MARDAN_CLUSTER_01_Rural','MARDAN_CLUSTER_01_Urban',
                            'MIRPUR_CLUSTER_01_Rural','MIRPUR_CLUSTER_01_Urban',
                            'MOTORWAY_CLUSTER_01_Rural',
                            'MURREE_CLUSTER_01_Rural','MURREE_CLUSTER_01_Urban',
                            'MUZAFFARABAD_CLUSTER_01_Rural','MUZAFFARABAD_CLUSTER_01_Urban',
                            'NOWSHEHRA_CLUSTER_01_Rural','NOWSHEHRA_CLUSTER_01_Urban',
                            'PESHAWAR_CLUSTER_01_Rural','PESHAWAR_CLUSTER_01_Urban',
                            'PESHAWAR_CLUSTER_02_Rural','PESHAWAR_CLUSTER_02_Urban',
                            'PESHAWAR_CLUSTER_03_Rural','PESHAWAR_CLUSTER_03_Urban',
                            'PESHAWAR_CLUSTER_04_Rural','PESHAWAR_CLUSTER_04_Urban',
                            'RAWALPINDI_CLUSTER_01_Rural','RAWALPINDI_CLUSTER_01_Urban',
                            'RAWALPINDI_CLUSTER_02_Rural','RAWALPINDI_CLUSTER_02_Urban',
                            'RAWALPINDI_CLUSTER_03_Urban','RAWALPINDI_CLUSTER_04_Rural',
                            'RAWALPINDI_CLUSTER_04_Urban','RAWALPINDI_CLUSTER_05_Rural',
                            'RAWALPINDI_CLUSTER_05_Urban',
                            'SWABI_CLUSTER_01_Rural','SWABI_CLUSTER_01_Urban',
                            'SWAT_CLUSTER_01_Rural','SWAT_CLUSTER_01_Urban',
                            'SWAT_CLUSTER_02_Rural',
                            'TALAGANG_CLUSTER_01_Rural',
                            'TAXILA_CLUSTER_01_Rural','TAXILA_CLUSTER_01_Urban'], 'North'),
                            dict.fromkeys(['CHAMAN_CLUSTER_21_Rural',
                            'CHAMAN_CLUSTER_21_Urban',
                            'DADU_CLUSTER_15_Rural','DADU_CLUSTER_15_Urban',
                            'GAWADAR_CLUSTER_20_Rural','GAWADAR_CLUSTER_20_Urban',
                            'GHOTKI_CLUSTER_09_Rural','GHOTKI_CLUSTER_09_Urban',
                            'HYDERABAD_CLUSTER_01_RURAL','HYDERABAD_CLUSTER_01_URBAN',
                            'HYDERABAD_CLUSTER_02_RURAL','HYDERABAD_CLUSTER_02_URBAN',
                            'HYDERABAD_CLUSTER_03_RURAL','HYDERABAD_CLUSTER_03_URBAN',
                            'HYDERABAD_CLUSTER_04_RURAL','HYDERABAD_CLUSTER_04_URBAN',
                            'JACOBABAD_CLUSTER_12_Rural','JACOBABAD_CLUSTER_12_Urban',
                            'KARACHI_CLUSTER_01_RURAL','KARACHI_CLUSTER_01_URBAN',
                            'KARACHI_CLUSTER_02_RURAL','KARACHI_CLUSTER_02_URBAN',
                            'KARACHI_CLUSTER_03_URBAN','KARACHI_CLUSTER_04_URBAN',
                            'KARACHI_CLUSTER_05_RURAL','KARACHI_CLUSTER_05_URBAN',
                            'KARACHI_CLUSTER_06_URBAN','KARACHI_CLUSTER_07_RURAL',
                            'KARACHI_CLUSTER_07_URBAN','KARACHI_CLUSTER_08_URBAN',
                            'KARACHI_CLUSTER_09_URBAN','KARACHI_CLUSTER_10_RURAL',
                            'KARACHI_CLUSTER_10_URBAN','KARACHI_CLUSTER_11_URBAN',
                            'KARACHI_CLUSTER_12_URBAN','KARACHI_CLUSTER_13_URBAN',
                            'KARACHI_CLUSTER_14_RURAL','KARACHI_CLUSTER_14_URBAN',
                            'KARACHI_CLUSTER_15_URBAN', 'KARACHI_CLUSTER_16_URBAN',
                            'KARACHI_CLUSTER_17_URBAN','KARACHI_CLUSTER_18_URBAN',
                            'KARACHI_CLUSTER_19_URBAN','KARACHI_CLUSTER_20_URBAN',
                            'KARACHI_CLUSTER_21_RURAL','KARACHI_CLUSTER_21_URBAN',
                            'KARACHI_CLUSTER_22_RURAL','KARACHI_CLUSTER_22_URBAN',
                            'KHUZDAR_CLUSTER_17_Rural','KHUZDAR_CLUSTER_17_Urban',
                            'LARKANA_CLUSTER_13_Rural','LARKANA_CLUSTER_13_Urban',
                            'MIRPURKHAS_CLUSTER_05_Rural','MIRPURKHAS_CLUSTER_05_Urban',
                            'MITHI_CLUSTER_04_Rural','MITHI_CLUSTER_04_Urban',
                            'MORO_CLUSTER_02_Rural','MORO_CLUSTER_02_Urban',
                            'NAWABSHAH_CLUSTER_01_Rural','NAWABSHAH_CLUSTER_01_Urban',
                            'QUETTAEAST_CLUSTER_19_Urban','QUETTAWEST_CLUSTER_18_Rural',
                            'QUETTAWEST_CLUSTER_18_Urban',
                            'SAKRAND_CLUSTER_06_Rural','SAKRAND_CLUSTER_06_Urban',
                            'SHAHDADKOT_CLUSTER_14_Rural','SHAHDADKOT_CLUSTER_14_Urban',
                            'SIBBI_CLUSTER_16_Rural', 'SIBBI_CLUSTER_16_Urban',
                            'SUI_CLUSTER_11_Rural','SUI_CLUSTER_11_Urban',
                            'SUKKUR_CLUSTER_10_Rural','SUKKUR_CLUSTER_10_Urban',
                            'THATTA_CLUSTER_07_Rural', 'THATTA_CLUSTER_07_Urban',
                            'UMERKOT_CLUSTER_03_Rural','UMERKOT_CLUSTER_03_Urban'], 'South'))

# apply mapping 
df1.loc[:, 'Region'] = df1['GCell Group'].map(df2.get)
```

## Calculate Counters Month Level


```python
def group_wa(series):
    return np.average(series.dropna())
```


```python
# Calculate intervals
df3 = df1.groupby(['GCell Group','Month','Region']).apply(lambda x: pd.Series({
        '_CallSetup TCH GOS(%)_N': (x['_CallSetup TCH GOS(%)_N'].sum()),
        '_CallSetup TCH GOS(%)_D': (x['_CallSetup TCH GOS(%)_D'].sum()),
        '_Mobility TCH GOS(%)_N': (x['_Mobility TCH GOS(%)_N'].sum()),
        '_Mobility TCH GOS(%)_D': (x['_Mobility TCH GOS(%)_D'].sum()),    
        '_GOS-SDCCH(%)_N': (x['_GOS-SDCCH(%)_N'].sum()),
        '_GOS-SDCCH(%)_D': (x['_GOS-SDCCH(%)_D'].sum()),   
        'CSSR_Non Blocking_1_N': (x['CSSR_Non Blocking_1_N'].sum()),
        'CSSR_Non Blocking_1_D': (x['CSSR_Non Blocking_1_D'].sum()),
        'CSSR_Non Blocking_2_N': (x['CSSR_Non Blocking_2_N'].sum()),
        'CSSR_Non Blocking_2_D': (x['CSSR_Non Blocking_2_D'].sum()),
        '_DCR_N': (x['_DCR_N'].sum()),
        '_DCR_D': (x['_DCR_D'].sum()),     
        'N_OutHSR_N': (x['N_OutHSR_N'].sum()),
        'N_OutHSR_D': (x['N_OutHSR_D'].sum()),
        '_RxQual Index UL_1': (x['_RxQual Index UL_1'].sum()),
        '_RxQual Index UL_2': (x['_RxQual Index UL_2'].sum()),
        '_RxQual Index DL_1': (x['_RxQual Index DL_1'].sum()),
        '_RxQual Index DL_2': (x['_RxQual Index DL_2'].sum()),
        'CS_Traffic': group_wa(x['GlobelTraffic']),
        'TCH Availability Rate': group_wa(x['TCH Availability Rate(%)'])
    })).reset_index()
```

## Calculate Month Level KPIs


```python
# GOS SDCCH
df3['SDCCH GoS']=(df3['_GOS-SDCCH(%)_N']/df3['_GOS-SDCCH(%)_D'])*100
```


```python
#TCH GOS
df3['TCH GoS']=(df3['_CallSetup TCH GOS(%)_N']/df3['_CallSetup TCH GOS(%)_D'])*100
```


```python
# MoB TCH GoS
df3['MoB GoS']=(df3['_Mobility TCH GOS(%)_N']/df3['_Mobility TCH GOS(%)_D'])*100
```


```python
# CSSR Non Blocking
df3['CSSR']=(1-(df3['CSSR_Non Blocking_1_N']/
                    df3['CSSR_Non Blocking_1_D']))*\
             (1-(df3['CSSR_Non Blocking_2_N']/
                 df3['CSSR_Non Blocking_2_D']))*100
```


```python
# DCR
df3['DCR']=(df3['_DCR_N']/df3['_DCR_D'])*100
```


```python
# Handover Success Rate
df3['HSR']=(df3['N_OutHSR_N']/df3['N_OutHSR_D'])*100
```


```python
# DL Quality 
df3['DL RQI']=(df3['_RxQual Index DL_1']/df3['_RxQual Index DL_2'])*100
```


```python
# UL Quality
df3['UL RQI']=(df3['_RxQual Index UL_1']/df3['_RxQual Index UL_2'])*100
```

## Filter Required Columns


```python
df4 = df3[['GCell Group','Month','Region','CSSR','DCR', 'HSR', 'DL RQI', 'UL RQI','SDCCH GoS', 'TCH GoS', 'MoB GoS',
            'CS_Traffic','TCH Availability Rate']]
```

## Filter only Clusters


```python
# Filter only cluster , remove city, region, sub region
df5=df4[df4['GCell Group']\
        .str.contains('|'.join(['_Rural','_Urban','RURAL','_URBAN']))].copy().reset_index(drop=True)
```

## Modification Cluster Name as per requirement


```python
# Cluster is Urban or Rural, get from GCell Group
df5['Cluster Type'] = df5['GCell Group'].str.strip().str[-5:]
# Remove Urban and Rural from GCell Group
df5['Location']=df5['GCell Group'].map(lambda x: str(x)[:-6])
```


```python
df5.loc[:, 'Cluster Type'] = df5['Cluster Type'].str.capitalize()
```

# Re-Shape (Pivot table)


```python
# re-shape()
df6=df5.pivot_table(index=['Region','Location','Month'],\
        columns='Cluster Type',\
        values=['CSSR','DCR','HSR',\
        'SDCCH GoS','TCH GoS','MoB GoS','DL RQI','UL RQI'],\
        aggfunc='mean').\
        sort_index(level=[0,1],axis=1,ascending=[True,False])
```

## Change the Column Sequence


```python
#Rrequired sequence
df7=df6[[('CSSR', 'Urban'),('CSSR', 'Rural'),\
           ('DCR', 'Urban'),('DCR', 'Rural'),\
            ('HSR', 'Urban'),('HSR', 'Rural'),\
            ('SDCCH GoS', 'Urban'),('SDCCH GoS', 'Rural'),\
            ('TCH GoS', 'Urban'),('TCH GoS', 'Rural'),\
            ('MoB GoS', 'Urban'),('MoB GoS', 'Rural'),\
            ('DL RQI', 'Urban'),('DL RQI', 'Rural'),\
            ('UL RQI', 'Urban'),('UL RQI', 'Rural')]].reset_index()

```

## Export Data Set


```python
with pd.ExcelWriter('SLA_Target_Month.xlsx',date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
        
    # Cluster BH Quater Level
    df4.to_excel(writer,sheet_name="BH_KPIs_Month",engine='openpyxl',na_rep='N/A',index=False)
    df7.to_excel(writer,sheet_name="BH_KPIs_Month_Format",engine='openpyxl',na_rep='N/A')
```


```python
# re-set all the variable from the RAM
%reset -f
```

# Formatting 


```python
import os
folder_path = 'D:/Advance_Data_Sets/SLA/Cluster_BH_KPIs'
os.chdir(folder_path)
```


```python
import openpyxl
from openpyxl import load_workbook

# Load the workbook
wb = load_workbook('SLA_Target_Month.xlsx')
```

## Zoom Scale Setting (All the sheets)


```python
# Zoom Scale Setting (All the sheets)
for ws in wb:
    ws.sheet_view.zoomScale = 80
```

## Apply border (All the Sheets)


```python
# Define the border style
border = openpyxl.styles.borders.Border(
    left=openpyxl.styles.borders.Side(style='thin'),
    right=openpyxl.styles.borders.Side(style='thin'),
    top=openpyxl.styles.borders.Side(style='thin'),
    bottom=openpyxl.styles.borders.Side(style='thin'))

# Loop through each cell in all the worksheet and set the border style
for ws in wb:
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
```

## Apply Font Style and Size (All the Sheets)


```python
# set the Font style and size of the Font of all the sheets
for ws in wb:
    for row in ws.iter_rows():
        for cell in row:
            cell.font = openpyxl.styles.\
            Font(name='Calibri Light', size=11)
```

## Apply alignment (All the sheets)


```python
for ws in wb:
    for row in ws.iter_rows():
        for cell in row:
             cell.alignment = openpyxl.styles.Alignment(horizontal='center',\
                                                        vertical='center')
```

## Set Float Number Format (For All the tabs)


```python
for ws in wb:
    for row in ws.iter_rows(min_col=1, max_col=ws.max_column):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = '0.00'
```

## Set Tab Color (All the Tabs)


```python
colors = ["00B0F0", "FFFF00", "00FF00", "FF00FF"]
for i, ws in enumerate(wb):
    ws.sheet_properties.tabColor = colors[i % len(colors)]
```

## Formatting 'BH_KPIs_Month_Format' Sheet


```python
# Select the worksheet
ws1 = wb['BH_KPIs_Month_Format']
```

### Delete the Row


```python
ws1.delete_rows(3)
```

### Merge Row


```python
# Merge cells in rows B1:B2
ws1.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)
# Merge cells in rows C1:C2
ws1.merge_cells(start_row=1, start_column=3, end_row=2, end_column=3)
# Merge cells in rows D1:D2
ws1.merge_cells(start_row=1, start_column=4, end_row=2, end_column=4)
```

### Hide Column


```python
# hide the 1st column
ws1.column_dimensions.group('A', hidden=True)
```

## Set Column width


```python
from openpyxl.utils import get_column_letter

# Set the width of column B to 12
ws1.column_dimensions[get_column_letter(2)].width = 12

# Set the width of column C to 25
ws1.column_dimensions[get_column_letter(3)].width = 25

# Set the width of column D to 12
ws1.column_dimensions[get_column_letter(4)].width = 12


# Set the width of all other columns to 8
for col in range(6,ws1.max_column+1):
    ws1.column_dimensions[get_column_letter(col)].width = 8
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


```python
# Define range_string
last_row = ws1.max_row
```

### 1. CSSR Urban


```python
range_string_E = "E3:E" + str(last_row)   # CSSR-Urban

# Define the conditional formatting rules for E column (CSSR URBAN)
ws1.conditional_formatting.add(range_string_E, 
                               CellIsRule(operator='lessThanOrEqual', 
                               formula=[99.3], 
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws1.conditional_formatting.add(range_string_E, 
                               CellIsRule(operator='between', 
                               formula=[99.3,99.5], 
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_E, 
                               CellIsRule(operator='greaterThanOrEqual', 
                                formula=[99.5],
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 2. CSSR Rural


```python
range_string_F = "F3:F" + str(last_row)   # CSSR-Rural
# Define the conditional formatting rules for F column (CSSR_RURAL)
ws1.conditional_formatting.add(range_string_F, 
                               CellIsRule(operator='lessThanOrEqual', 
                               formula=[98.7], 
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws1.conditional_formatting.add(range_string_F, 
                               CellIsRule(operator='between', 
                                          formula=[98.7,99.0], 
                                          fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_F, 
                               CellIsRule(operator='greaterThanOrEqual', 
                               formula=[99.0], 
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 3. DCR Urban


```python
range_string_G = "G3:G" + str(last_row)   # DCR-Urban

# Define the conditional formatting rules for F column (DCR-Urban)
ws1.conditional_formatting.add(range_string_G, \
                            CellIsRule(operator='lessThanOrEqual', \
                            formula=[0.6], \
                            fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws1.conditional_formatting.add(range_string_G,\
                               CellIsRule(operator='between', \
                               formula=[0.6,0.65], \
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_G, \
                               CellIsRule(operator='greaterThanOrEqual', \
                                formula=[0.65], \
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 4. DCR Rural


```python
range_string_H = "H3:H" + str(last_row)   # DCR-Rural
# Define the conditional formatting rules for G column (DCR-Rural)
ws1.conditional_formatting.add(range_string_H,
                               CellIsRule(operator='lessThanOrEqual', 
                            formula=[1.0],
                            fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws1.conditional_formatting.add(range_string_H, 
                            CellIsRule(operator='between', 
                            formula=[1.0,1.05], 
                            fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_H, 
                               CellIsRule(operator='greaterThanOrEqual', 
                               formula=[1.05], 
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 5. HSR Urban 


```python
range_string_I = "I3:I" + str(last_row)   # HSR-Urban

# Define the conditional formatting rules for H column (HSR-Urban)
ws1.conditional_formatting.add(range_string_I, 
                               CellIsRule(operator='lessThanOrEqual', 
                                formula=[97], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws1.conditional_formatting.add(range_string_I, 
                            CellIsRule(operator='between', 
                            formula=[97,97.5], 
                            fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_I, 
                                CellIsRule(operator='greaterThanOrEqual', 
                                formula=[97.5], 
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 6. HSR Rural 


```python
range_string_J = "J3:J" + str(last_row)   # HSR-Rural
# Define the conditional formatting rules for I column (HSR-Rural)
ws1.conditional_formatting.add(range_string_J, 
                               CellIsRule(operator='lessThanOrEqual', 
                                formula=[95.6], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws1.conditional_formatting.add(range_string_J, 
                               CellIsRule(operator='between', formula=[95.6,96], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_J, 
                               CellIsRule(operator='greaterThanOrEqual', 
                               formula=[96.0], 
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 7. SDCCH GoS-Urban


```python
range_string_K = "K3:K" + str(last_row)   # SDCCH GoS-Urban

# Define the conditional formatting rules for J column (SDCCH GoS-Urban)
ws1.conditional_formatting.add(range_string_K, 
                               CellIsRule(operator='lessThanOrEqual', 
                               formula=[0.12], 
                               fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws1.conditional_formatting.add(range_string_K, 
                               CellIsRule(operator='between', 
                               formula=[0.1,0.12], 
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_K, 
                               CellIsRule(operator='greaterThanOrEqual', 
                                formula=[0.12], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 8. SDCCH GoS-Rural


```python
range_string_L = "L3:L" + str(last_row)   # SDCCH GoS-Rural
# Define the conditional formatting rules for K column (SDCCH GoS-Rural)
ws1.conditional_formatting.add(range_string_L, 
                               CellIsRule(operator='lessThanOrEqual', 
                               formula=[0.1], 
                               fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))


ws1.conditional_formatting.add(range_string_L, 
                               CellIsRule(operator='between', 
                                formula=[0.1,0.12], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))


ws1.conditional_formatting.add(range_string_L, 
                               CellIsRule(operator='greaterThanOrEqual', 
                                formula=[0.12], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 9.TCH GoS-Urban


```python
range_string_M = "M3:M" + str(last_row)   # TCH GoS-Urban

# Define the conditional formatting rules for L column (TCH GoS-Urban)
ws1.conditional_formatting.add(range_string_M, 
                               CellIsRule(operator='lessThanOrEqual', 
                                formula=[2.0], 
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws1.conditional_formatting.add(range_string_M, 
                               CellIsRule(operator='between', 
                                formula=[2.,2.2], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_M, 
                               CellIsRule(operator='greaterThanOrEqual', 
                                formula=[2.2], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 10.TCH GoS-Rural


```python
range_string_N = "N3:N" + str(last_row)   # TCH GoS-Rural
# Define the conditional formatting rules for M column (TCH GoS-Rural)
ws1.conditional_formatting.add(range_string_N, 
                               CellIsRule(operator='lessThanOrEqual', 
                                          formula=[2.0], 
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws1.conditional_formatting.add(range_string_N, 
                               CellIsRule(operator='between', 
                                formula=[2,2.2], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_N, 
                               CellIsRule(operator='greaterThanOrEqual', 
                                formula=[2.2], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 11. MoB GoS-Urban


```python
range_string_O = "O3:O" + str(last_row)   # MoB GoS-Urban

# Define the conditional formatting rules for N column (MoB GoS-Urban)
ws1.conditional_formatting.add(range_string_O, 
                               CellIsRule(operator='lessThanOrEqual', 
                                formula=[4.0], 
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws1.conditional_formatting.add(range_string_O, 
                               CellIsRule(operator='between', 
                                formula=[4.0,4.2], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_O, 
                               CellIsRule(operator='greaterThanOrEqual',
                                formula=[4.2], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 12. MoB GoS-Rural


```python
range_string_P = "P3:P" + str(last_row)   # MoB GoS-Rural
# Define the conditional formatting rules for O column (MoB GoS-Rural)
ws1.conditional_formatting.add(range_string_P, 
                               CellIsRule(operator='lessThanOrEqual', 
                                formula=[4.0], 
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))

ws1.conditional_formatting.add(range_string_P, 
                               CellIsRule(operator='between', 
                                formula=[4.0,4.2], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_P, 
                               CellIsRule(operator='greaterThanOrEqual',
                                formula=[4.2], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))
```

### 13. DL RQI-Urban


```python
range_string_Q = "Q3:Q" + str(last_row)   # DL RQI-Urban

# Define the conditional formatting rules for P column (DL RQI-Urban)
ws1.conditional_formatting.add(range_string_Q, 
                               CellIsRule(operator='lessThanOrEqual',
                               formula=[98.0], 
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws1.conditional_formatting.add(range_string_Q, 
                               CellIsRule(operator='between', 
                                formula=[98.0,98.4],
                                fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_Q,
                               CellIsRule(operator='greaterThanOrEqual', 
                                formula=[98.4],
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 14. DL RQI-Rural


```python
range_string_R = "R3:R" + str(last_row)   # DL RQI-Rural
# Define the conditional formatting rules for Q column (DL RQI-Rural)
ws1.conditional_formatting.add(range_string_R, 
                               CellIsRule(operator='lessThanOrEqual', 
                               formula=[96.5], 
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws1.conditional_formatting.add(range_string_R, 
                               CellIsRule(operator='between', 
                               formula=[96.5,97], 
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_R,
                               CellIsRule(operator='greaterThanOrEqual', 
                                formula=[97.0], 
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 15. UL RQI-Urban


```python
range_string_S = "S3:S" + str(last_row)   # UL RQI-Urban

# Define the conditional formatting rules for R column (UL RQI-Urban)
ws1.conditional_formatting.add(range_string_S, 
                               CellIsRule(operator='lessThanOrEqual', 
                                formula=[98.0], 
                                fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws1.conditional_formatting.add(range_string_S, 
                               CellIsRule(operator='between', 
                               formula=[98.0,98.2], 
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_S, 
                               CellIsRule(operator='greaterThanOrEqual', 
                               formula=[98.2], 
                               fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

### 16. UL RQI-Rural


```python
range_string_T = "T3:T" + str(last_row)   # UL RQI-Rural
# Define the conditional formatting rules for S column (UL RQI-Rural)
ws1.conditional_formatting.add(range_string_T,
                               CellIsRule(operator='lessThanOrEqual',
                               formula=[97.2], 
                               fill=openpyxl.styles.PatternFill(start_color='FFFF0000', end_color='FFFF0000')))

ws1.conditional_formatting.add(range_string_T, 
                               CellIsRule(operator='between', 
                               formula=[97.2,97.7], 
                               fill=openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00')))

ws1.conditional_formatting.add(range_string_T, 
                               CellIsRule(operator='greaterThanOrEqual', 
                               formula=[97.7], 
                               fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00')))
```

## Header Format


```python
# set background and font color
for row in range(1,3):
    for cell in ws1[row]:
        cell.fill = openpyxl.styles.PatternFill(start_color="C00000", end_color="C00000", fill_type = "solid")
        font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=11,name='Calibri Light')
        cell.font = font
```

## Formatting 'BH_KPIs_Month_Format' Sheet


```python
# Select the worksheet
ws2 = wb['BH_KPIs_Month']
```

## Column Width


```python
# Iterate over all columns in the sheet
for column in ws2.columns:
    # Get the current width of the column
    current_width = ws2.column_dimensions[column[0].column_letter].width
    # Get the maximum width of the cells in the column
    length = max(len(str(cell.value)) for cell in column)
    # Set the width of the column to fit the maximum width, if it's greater than the current width
    if length > current_width:
            ws2.column_dimensions[column[0].column_letter].width = length     
```

## Header Formatting


```python
for cell in ws2[1]:
    cell.fill = openpyxl.styles.PatternFill(start_color="0070C0", end_color="0070C0", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True)
    cell.font = font
```

## Insert a New Sheet (as First Sheet)


```python
ws511 =wb.create_sheet("Title Page", 0)
```

## Merge Specific Row and Columns


```python
ws511.merge_cells(start_row=12, start_column=5, end_row=18, end_column=18)
```

## Fill the Merge Cells


```python
# Fill the text 'CE Utilization Report' in the merged cell
ws511.cell(row=12, column=5).value = 'Cluster BH KPIs Month Level'
```

## Formatting Tital Page Report


```python
# Access the first row starting from row 3
first_row1 = list(ws511.rows)[11]
# Iterate through the cells in the first row starting from column E
for cell in first_row1[4:]:
    cell.alignment = openpyxl.styles.Alignment(horizontal='center', \
                                            vertical='center')
    cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", \
                                            end_color="CF0A2C", \
                                            fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",\
                                            bold=True,\
                                            size=65,\
                                            name='Calibri Light')
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
from openpyxl.utils import get_column_letter
'''loop through all sheets in the workbook and 
    insert the hyperlink to each sheet
'''
row = 22
for ws in wb:
    if ws.title != "Title Page":
        hyperlink_cell = ws511.cell(row=row, column=5)
        hyperlink_cell.value = ws.title
        hyperlink_cell.hyperlink = "#'{}'!A1".format(ws.title)
        hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", \
                                                underline="single")
        hyperlink_cell.border = border
        hyperlink_cell.alignment = openpyxl.styles.\
                        Alignment(horizontal='center')
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
        hyperlink_cell.font = openpyxl.styles.Font(color="0000FF",\
                                                underline="single")
        hyperlink_cell.alignment = openpyxl.styles.\
                                Alignment(horizontal='center')
        
        # Set border on the hyperlink column
        for row in ws.iter_rows(min_row=2, \
                                max_row=4, \
                                min_col=hyperlink_cell.column, \
                                max_col=hyperlink_cell.column):
            for cell in row:
                cell.border = border
        # Set width of the hyperlink column
        col_letter = openpyxl.utils.get_column_letter(hyperlink_cell.column)
        ws.column_dimensions[col_letter].width = 25
        
        '''
        # Add hyperlink to cell in the last 
        column+2 of the sheet for next sheet
        '''
        if i < len(wb.worksheets)-1:
            next_hyperlink_cell = ws.cell(row=3, column=ws.max_column)
            next_hyperlink_cell.value = "Next Sheet"
            next_hyperlink_cell.hyperlink = "#'{}'!A1".\
                        format(wb.worksheets[i+1].title)
            next_hyperlink_cell.font = openpyxl.styles.Font(color="0000FF",\
                                                        underline="single")
            next_hyperlink_cell.alignment = openpyxl.styles.\
                                Alignment(horizontal='center')
        '''
        # Add hyperlink to cell in the last 
        column+2 of the sheet for previous sheet
        '''
        if i > 0:
            prev_hyperlink_cell = ws.cell(row=4, column=ws.max_column)
            prev_hyperlink_cell.value = "Previous Sheet"
            prev_hyperlink_cell.hyperlink = "#'{}'!A1".format(wb.worksheets[i-1].title)
            prev_hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", \
                                                        underline="single")
            prev_hyperlink_cell.alignment = openpyxl.styles.\
                                Alignment(horizontal='center')
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

## Final Output


```python
# Save the changes
wb.save('SLA_Target_Month.xlsx')
```

## Re-Set Variables


```python
#re-set all the variable from the RAM
%reset -f
```


```python
import shutil

# set the file path and folder paths
file_path = "D:/Advance_Data_Sets/SLA/Cluster_BH_KPIs/SLA_Target_Month.xlsx"
destination_folder = "D:/Advance_Data_Sets/Output_Folder"

# use the shutil.move() function to move the file
shutil.move(file_path, destination_folder)
```
