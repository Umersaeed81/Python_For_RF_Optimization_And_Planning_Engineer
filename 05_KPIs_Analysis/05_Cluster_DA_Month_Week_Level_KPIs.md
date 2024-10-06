#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# Monthly & Weekly DA KPIs
tt
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
folder_path = 'D:/Advance_Data_Sets/SLA/Cluster_DA_KPIs'
os.chdir(folder_path)
```

## Import Cluster DA Counters


```python
df = sorted(glob('*.zip'))
# concat all the Cluster DA Files
df0=pd.concat((pd.read_csv(file,skiprows=range(6),encoding='unicode_escape',
        skipfooter=1,engine='python',
        usecols=['Date','GCell Group',
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

## Assign Week Number


```python
df0['Week']=df0['Date'].dt.isocalendar().week.astype(str).apply(lambda x: x.zfill(2))\
                                                            +"_"+ \
                                        df0['Date'].dt.isocalendar().year.astype(str)
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

## Calculate Counters Week Level


```python
# Calculate intervals
df4 = df1.groupby(['GCell Group','Week','Region']).apply(lambda x: pd.Series({
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


```python
# Filter Required Columns from Monthly KPIs
df5 = df3[['GCell Group','Month','Region','CSSR','DCR', 'HSR', 'DL RQI', 'UL RQI','SDCCH GoS', 'TCH GoS', 'MoB GoS',
            'CS_Traffic','TCH Availability Rate']]
```

## Calculate Week Level KPIs


```python
# GOS SDCCH
df4['SDCCH GoS']=(df4['_GOS-SDCCH(%)_N']/df4['_GOS-SDCCH(%)_D'])*100
```


```python
#TCH GOS
df4['TCH GoS']=(df4['_CallSetup TCH GOS(%)_N']/df4['_CallSetup TCH GOS(%)_D'])*100
```


```python
# MoB TCH GoS
df4['MoB GoS']=(df4['_Mobility TCH GOS(%)_N']/df4['_Mobility TCH GOS(%)_D'])*100
```


```python
# CSSR Non Blocking
df4['CSSR']=(1-(df4['CSSR_Non Blocking_1_N']/
                    df4['CSSR_Non Blocking_1_D']))*\
             (1-(df4['CSSR_Non Blocking_2_N']/
                 df4['CSSR_Non Blocking_2_D']))*100
```


```python
# DCR
df4['DCR']=(df4['_DCR_N']/df4['_DCR_D'])*100
```


```python
# Handover Success Rate
df4['HSR']=(df4['N_OutHSR_N']/df4['N_OutHSR_D'])*100
```


```python
# DL Quality 
df4['DL RQI']=(df4['_RxQual Index DL_1']/df4['_RxQual Index DL_2'])*100
```


```python
# UL Quality
df4['UL RQI']=(df4['_RxQual Index UL_1']/df4['_RxQual Index UL_2'])*100
```


```python
# Filter Required Columns from Weekly KPIs
df6 = df4[['GCell Group','Week','Region','CSSR','DCR', 'HSR', 'DL RQI', 'UL RQI','SDCCH GoS', 'TCH GoS', 'MoB GoS',
            'CS_Traffic','TCH Availability Rate']]
```

## Export Output DataFrame


```python
with pd.ExcelWriter('Cluster_DA_KPIs.xlsx',date_format = 'dd-mm-yyyy',datetime_format='dd-mm-yyyy') as writer:
    df5.to_excel(writer,sheet_name="DA_KPIs_Month",engine='openpyxl',na_rep='N/A',index=False)
    df6.to_excel(writer,sheet_name="DA_KPIs_Week",engine='openpyxl',na_rep='N/A',index=False)
```


```python
#re-set all the variable from the RAM
%reset -f
```

# Formatting 

## Import required Libraries


```python
import os
import openpyxl
from openpyxl import load_workbook
```

## Set Input File Path


```python
folder_path = 'D:/Advance_Data_Sets/SLA/Cluster_DA_KPIs'
os.chdir(folder_path)
```

## Load Input File Path


```python
# Load the workbook
wb = load_workbook('Cluster_DA_KPIs.xlsx')
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
             cell.alignment = openpyxl.styles.Alignment(horizontal='center',vertical='center')
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

## WrapText of header and Formatting (All the sheets)


```python
# wrape text of the header and center aligment
for ws in wb:
    for row in ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrapText=True, \
                                                    horizontal='center',\
                                                    vertical='center')
            cell.fill = openpyxl.styles.PatternFill(start_color="C00000",\
                                                    end_color="C00000", \
                                                    fill_type = "solid")
            font = openpyxl.styles.Font(color="FFFFFF",\
                                                    bold=True,\
                                                    size=11,\
                                                    name='Calibri Light')
            cell.font = font
```

## Column Width (All the Tabs)


```python
# Iterate over all sheets
for ws in wb.worksheets:
    # Iterate over all columns in the sheet
    for column in ws.columns:
        # Get the current width of the column
        current_width = ws.column_dimensions[column[0].column_letter].width
        # Get the maximum width of the cells in the column
        length = max(len(str(cell.value)) for cell in column)
        # Set the width of the column to fit the maximum width, if it's greater than the current width
        if length > current_width:
            ws.column_dimensions[column[0].column_letter].width = length
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
ws511.cell(row=12, column=5).value = 'Cluster DA KPIs'
```

## Formatting Tital Page Report


```python
# Access the first row starting from row 3
first_row1 = list(ws511.rows)[11]
# Iterate through the cells in the first row starting from column E
for cell in first_row1[4:]:
    cell.alignment = openpyxl.styles.Alignment(horizontal='center',  vertical='center')
    cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", end_color="CF0A2C", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=70, name='Calibri Light')
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
wb.save('Cluster_DA_KPIs.xlsx')
```

## Re-Set Variables


```python
#re-set all the variable from the RAM
%reset -f
```


```python
import shutil

# set the file path and folder paths
file_path = "D:/Advance_Data_Sets/SLA/Cluster_DA_KPIs/Cluster_DA_KPIs.xlsx"
destination_folder = "D:/Advance_Data_Sets/Output_Folder"

# use the shutil.move() function to move the file
shutil.move(file_path, destination_folder)
```
