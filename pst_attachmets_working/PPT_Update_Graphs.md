## 📚 Import Required Libraries


```python
import os
import pandas as pd
from pptx.util import Inches
from pptx import Presentation
from pptx.chart.data import CategoryChartData
#!pip install python-pptx
```

## 📁 Set Input File Path


```python
path = 'D:/Advance_Data_Sets/UMTS_SunSet/PPT/City_Lelvel_Traffic'
os.chdir(path)
```

## 🔍 Inspecting Shapes and Identifying Charts on a Specific Slide in PowerPoint


```python
# path = 'D:/Advance_Data_Sets/UMTS_SunSet/PPT/PPT_Temp'
# os.chdir(path)

# from pptx import Presentation

# # Load PowerPoint file
# prs = Presentation('template.pptx')

# # Choose the slide number (e.g., second slide)
# slide = prs.slides[2]  # 0-based index

# # Loop through and print info about each shape
# for idx, shape in enumerate(slide.shapes):
#     shape_type = shape.shape_type
#     name = getattr(shape, "name", "No name")
#     print(f"Index: {idx}, Type: {shape_type}, Name: {name}, Has Chart: {shape.has_chart}")
```

### 1) Bhakar - Data_Volume (All Technology)


```python
df = pd.read_excel('Bhakar_Target_City_Metrics_Report.xlsx',sheet_name='Data_Volume').\
            rename(columns={'Date': 'Row Labels',
                            '3G_DataVolume_GB': '3G Data Volume (GB)',
                            '4G_DataVolume_GB':'4G Data Volume (GB)',
                           'Total_Data_volume':'Total Data Volume (GB)'})
```

### 2) DI Khan - Data Volume (All Technology)


```python
df1 = pd.read_excel('DI_ Khan_Control_City_Metrics_Report.xlsx',sheet_name='Data_Volume').\
            rename(columns={'Date': 'Row Labels',
                            '3G_DataVolume_GB': '3G Data Volume (GB)',
                            '4G_DataVolume_GB':'4G Data Volume (GB)',
                           'Total_Data_volume':'Total Data Volume (GB)'})
```

### 3) Bhaker - Data Volume (LTE)


```python
df2 = pd.read_excel('4G_Data_Volume_Report.xlsx',sheet_name='4G_Bhakar',usecols=['Date','L1800','L2100','L900','City']).\
            rename(columns={'Date': 'Row Labels',
                            'L1800':'L1800',
                            'L2100':'L2100',
                            'L900':'L900',
                            'City':'Total'}).fillna(0)
```

### 4) DI Khan - Data Volume (LTE)


```python
df3 = pd.read_excel('4G_Data_Volume_Report.xlsx',sheet_name='4G_DI_ Khan',usecols=['Date','L1800','L2100','City']).\
            rename(columns={'Date': 'Row Labels',
                            'L1800':'L1800',
                            'L2100':'L2100',
                            'City':'Total'})
```

### 5) Bhaker - Throughput


```python
df4 = pd.read_excel('4G_thoroughput_Report.xlsx',sheet_name='4G_Bhakar',usecols=['Date','City','L1800','L2100','L900']).\
                            rename(columns={'Date': 'Row Labels',
                            'City':'Overall',
                            'L1800':'L1800',
                            'L2100':'L2100',
                            'L900':'L900'}).fillna(0)
```

### 6) DI Khan - Throughput


```python
df5 = pd.read_excel('4G_thoroughput_Report.xlsx',sheet_name='4G_DI_ Khan',usecols=['Date','City','L1800','L2100']).\
                            rename(columns={'Date': 'Row Labels',
                            'City':'Overall',
                            'L1800':'L1800',
                            'L2100':'L2100',
                            'L900':'L900'})
```

### 7) Bhakar - CS Traffic


```python
df6 = pd.read_excel('Bhakar_Target_City_Metrics_Report.xlsx',sheet_name='CS_Traffic').\
            rename(columns={'Date': 'Row Labels',
                            '2G_CS_Traffic': '2G Voice',
                            '3G_CS_Traffic'	:'3G Voice',
                            '4G_CS_Traffic': 'VoLTE',
                           'Total_CS_Traffic':'Total Voice'})
```

### 8) DI Khan - CS Traffic


```python
df7 = pd.read_excel('DI_ Khan_Control_City_Metrics_Report.xlsx',sheet_name='CS_Traffic').\
            rename(columns={'Date': 'Row Labels',
                            '2G_CS_Traffic': '2G Voice',
                            '3G_CS_Traffic'	:'3G Voice',
                            '4G_CS_Traffic': 'VoLTE',
                           'Total_CS_Traffic':'Total Voice'})
```

### 9) Layyah - Data_Volume (All Technology)


```python
df8 = pd.read_excel('Layyah_Target_City_Metrics_Report.xlsx',sheet_name='Data_Volume').\
            rename(columns={'Date': 'Row Labels',
                            '3G_DataVolume_GB': '3G Data Volume (GB)',
                            '4G_DataVolume_GB':'4G Data Volume (GB)',
                           'Total_Data_volume':'Total Data Volume (GB)'})
```

### 10) Kot_Adu - Data_Volume (All Technology)


```python
df9 = pd.read_excel('Kot_Adu_Control_City_Metrics_Report.xlsx',sheet_name='Data_Volume').\
            rename(columns={'Date': 'Row Labels',
                            '3G_DataVolume_GB': '3G Data Volume (GB)',
                            '4G_DataVolume_GB':'4G Data Volume (GB)',
                           'Total_Data_volume':'Total Data Volume (GB)'})
```

### 11) Layyah - Data Volume (LTE)


```python
df10 = pd.read_excel('4G_Data_Volume_Report.xlsx',sheet_name='4G_Layyah',usecols=['Date','L1800','L2100','L900','City']).\
            rename(columns={'Date': 'Row Labels',
                            'L1800':'L1800',
                            'L2100':'L2100',
                            'L900':'L900',
                            'City':'Total'}).fillna(0)
```

### 12) Kot_Adu - Data Volume (LTE)


```python
df11 = pd.read_excel('4G_Data_Volume_Report.xlsx',sheet_name='4G_Kot_Adu',usecols=['Date','L1800','L2100','City']).\
            rename(columns={'Date': 'Row Labels',
                            'L1800':'L1800',
                            'L2100':'L2100',
                            'City':'Total'})
```

### 13) Layyah - Throughput


```python
df12 = pd.read_excel('4G_thoroughput_Report.xlsx',sheet_name='4G_Layyah',usecols=['Date','City','L1800','L2100','L900']).\
                            rename(columns={'Date': 'Row Labels',
                            'City':'Overall',
                            'L1800':'L1800',
                            'L2100':'L2100',
                            'L900':'L900'}).fillna(0)
```

### 11) Kot_Adu - Throughput


```python
df13 = pd.read_excel('4G_thoroughput_Report.xlsx',sheet_name='4G_Kot_Adu',usecols=['Date','City','L1800','L2100']).\
                            rename(columns={'Date': 'Row Labels',
                            'City':'Overall',
                            'L1800':'L1800',
                            'L2100':'L2100',
                            'L900':'L900'})
```

### 12) Layyah - CS Traffic


```python
df14 = pd.read_excel('Layyah_Target_City_Metrics_Report.xlsx',sheet_name='CS_Traffic').\
            rename(columns={'Date': 'Row Labels',
                            '2G_CS_Traffic': '2G Voice',
                            '3G_CS_Traffic'	:'3G Voice',
                            '4G_CS_Traffic': 'VoLTE',
                           'Total_CS_Traffic':'Total Voice'})
```

### 13) Kot_Adu - CS Traffic


```python
df15 = pd.read_excel('Kot_Adu_Control_City_Metrics_Report.xlsx',sheet_name='CS_Traffic').\
            rename(columns={'Date': 'Row Labels',
                            '2G_CS_Traffic': '2G Voice',
                            '3G_CS_Traffic'	:'3G Voice',
                            '4G_CS_Traffic': 'VoLTE',
                           'Total_CS_Traffic':'Total Voice'})
```

## 🧩 User-Defined Function to Update PowerPoint Chart with DataFrame


```python
def update_chart(prs, slide_index, shape_index, dataframe):
    chart_data = CategoryChartData()
    chart_data.categories = dataframe['Row Labels']

    for col in dataframe.columns[1:]:  # Skip 'Row Labels'
        chart_data.add_series(col, dataframe[col])

    slide = prs.slides[slide_index]
    chart_shape = slide.shapes[shape_index]
    chart = chart_shape.chart
    chart.replace_data(chart_data)
```

## 🗂️📊 Load PowerPoint template for chart updates 


```python
path = 'D:/Advance_Data_Sets/UMTS_SunSet/PPT/PPT_Temp'
os.chdir(path)
# Load PowerPoint and update charts
prs = Presentation('template.pptx')
```

## 🔁 Reusing User-Defined Function to Update Slide Charts


```python
update_chart(prs, slide_index=1, shape_index=0, dataframe=df)          # Bhakar - Data Volume
update_chart(prs, slide_index=1, shape_index=3, dataframe=df1)         # DI Khan - Date Volume
update_chart(prs, slide_index=2, shape_index=2, dataframe=df2)         # Bhakar - Data Volume (LTE)
update_chart(prs, slide_index=2, shape_index=3, dataframe=df3)         # DI Khan - Data Volume (LTE)
update_chart(prs, slide_index=3, shape_index=0, dataframe=df4)         # Bhakar - Througput
update_chart(prs, slide_index=3, shape_index=3, dataframe=df5)         # DI Khan - Througput
update_chart(prs, slide_index=4, shape_index=0, dataframe=df6)         # Bhakar - CS Traffic
update_chart(prs, slide_index=4, shape_index=3, dataframe=df7)         # DI Khan - CS Traffic
update_chart(prs, slide_index=5, shape_index=0, dataframe=df8)         # Layyah - Data Volume
update_chart(prs, slide_index=5, shape_index=3, dataframe=df9)         # Kot_Adu - Data Volume
update_chart(prs, slide_index=6, shape_index=2, dataframe=df10)        # Layyah - Data Volume (LTE)
update_chart(prs, slide_index=6, shape_index=3, dataframe=df11)        # Kot_Adu - Data Volume (LTE)
update_chart(prs, slide_index=7, shape_index=0, dataframe=df12)        # Layyah - Througput
update_chart(prs, slide_index=7, shape_index=3, dataframe=df13)        # Kot_Adu - Througput
update_chart(prs, slide_index=8, shape_index=0, dataframe=df14)        # Layyah - CS Traffic
update_chart(prs, slide_index=8, shape_index=3, dataframe=df15)        # Kot_Adu - CS Traffic
```

## 💾 Save Presentation


```python
prs.save('updated_presentation.pptx')
```
