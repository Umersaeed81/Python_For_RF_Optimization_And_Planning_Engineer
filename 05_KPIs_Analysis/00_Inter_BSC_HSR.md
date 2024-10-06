#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# Inter BSC HSR

## Import Requied Libraries


```python
# Python libraries
import os
import pandas as pd
from glob import glob

```

## Set Input File Path


```python
# Set Working Folder Path For BSC Busy Hour KPIs
path= 'D:\Advance_Data_Sets\KPIs_Analysis\Inter_BSC_HSR\KPIs'
os.chdir(path)
```

## Import and Concat BSC Busy Hour Data 


```python
# List Zip Fil in a folder
df= sorted(glob('*.zip'))
# concat all the BSCs Busy Hour KPIs
df0=pd.concat((pd.read_csv(file,skiprows=range(6),\
                                skipfooter=1,engine='python',\
                                parse_dates=["Date"],na_values=['NIL','/0']) \
                                for file in df)).drop(["Integrity","Time"], axis=1)
```

## Assign Week Number to the Date 


```python
df0['WK']= df0['Date'].dt.isocalendar().year.astype(str) +"_"+ \
           df0['Date'].dt.isocalendar().week.astype(str).apply(lambda x: x.zfill(2))
```

## Assign Quater Number to the Date 


```python
# Identify the Quater 
df0['Quater'] = pd.PeriodIndex(pd.to_datetime(df0.Date), freq='Q')
```


```python
df0['Month']=  df0['Date'].dt.year.astype(str)+"_"+ \
               df0['Date'].dt.month.astype(str).apply(lambda x: x.zfill(2))
```

## BSC Defination File


```python
g= [('HLKNBSC02', 'Rural', 'South', 83), ('HSUKBSC04', 'Rural', 'South', 83), ('HLKNBSC03', 'Rural', 'South', 83),
    ('HDGKBSC05', 'Rural', 'Central 3', 83),('HQTABSC07', 'Rural', 'South', 83),('HCHKBSC03', 'Rural', 'North', 83),
    ('HDGKBSC04', 'Rural', 'Central 3', 83), ('HSUKBSC06', 'Rural', 'South', 83),('HISBBSC11', 'Rural', 'North', 83),
    ('HHYDBSC03', 'Rural', 'South', 83),('HDGKBSC03', 'Rural', 'Central 3', 83),('HPSWBSC04', 'Rural', 'North', 83),
    ('HDIKBSC03', 'Rural', 'Central 2', 83),('HSKTBSC04', 'Rural', 'Central 1', 83),('HPSWBSC03', 'Rural', 'North', 83),
    ('HHYDBSC04', 'Rural', 'South', 83), ('HKHIBSC17', 'Rural', 'South', 83),('HRYKBSC03', 'Rural', 'Central 3', 83),
    ('HQTABSC03', 'Rural', 'South', 83), ('HMDNBSC03', 'Rural', 'North', 83), ('HJHLBSC03', 'Rural', 'North', 83),
    ('HHFZBSC01', 'Rural', 'Central 1', 83), ('HBWPBSC01', 'Rural', 'Central 3', 83),('HGJWBSC04', 'Rural', 'Central 1', 83),
    ('HISBBSC20', 'Rural', 'North', 83), ('HSUKBSC05', 'Rural', 'South', 83),('HCHKBSC02', 'Rural', 'North', 83),
    ('HKHTBSC02', 'Rural', 'North', 83),('HSKTBSC03', 'Rural', 'Central 1', 83),('HGJTBSC06', 'Rural', 'Central 1', 83),
    ('HNWBBSC02', 'Rural', 'South', 83), ('HPSWBSC06', 'Rural', 'North', 83), ('HTTSBSC03', 'Rural', 'Central 2', 83),
    ('HMLTBSC06', 'Rural', 'Central 3', 83), ('HSGDBSC02', 'Rural', 'Central 2', 83), ('HQTABSC05', 'Rural', 'South', 83),
    ('HMPKBSC02', 'Rural', 'South', 83), ('HKHIBSC02', 'Rural', 'South', 83), ('HKASBSC03', 'Rural', 'Central 1', 83),
    ('HKHTBSC03', 'Rural', 'North', 83), ('HMDNBSC02', 'Rural', 'North', 83), ('HFSDBSC08', 'Rural', 'Central 2', 83),
    ('HISBBSC17', 'Rural', 'North', 83),('HKNWBSC01', 'Rural', 'Central 3', 83),('HMWLBSC03', 'Rural', 'Central 2', 83),
    ('HGJTBSC05', 'Rural', 'Central 1', 83), ('HSGDBSC03', 'Rural', 'Central 2', 83),('HLHRBSC11', 'Rural', 'Central 1', 83),
    ('HABTBSC02', 'Rural', 'North', 83), ('HHFZBSC02', 'Rural', 'Central 1', 83), ('HSWLBSC03', 'Rural', 'Central 2', 83),
    ('HMWLBSC04', 'Rural', 'Central 2', 83), ('HBWPBSC02', 'Rural', 'Central 3', 83), ('HKASBSC02', 'Rural', 'Central 1', 83),
    ('HISBBSC18', 'Rural', 'North', 83),('HKNWBSC02', 'Rural', 'Central 3', 83),('HFSDBSC07', 'Rural', 'Central 2', 83),
    ('HJHLBSC02', 'Rural', 'North', 83), ('HQTABSC08', 'Rural', 'South', 83),('HSWLBSC02', 'Rural', 'Central 2', 83),
    ('HDIKBSC02', 'Rural', 'Central 2', 83),('HMNSBSC03', 'Rural', 'North', 83),('HKHTBSC04', 'Rural', 'North', 83),
    ('HMLTBSC08', 'Rural', 'Central 3', 83),('HTTSBSC02', 'Rural', 'Central 2', 83),('HBWNBSC01', 'Rural', 'Central 3', 83),
    ('HBRWBSC03', 'Rural', 'Central 3', 83), ('HBWNBSC02', 'Rural', 'Central 3', 83),('HBRWBSC02', 'Rural', 'Central 3', 83),
    ('HQTABSC06', 'Urban', 'South', 90), ('HQTABSC04', 'Urban', 'South', 90),('HISBBSC19', 'Urban', 'North', 90),
    ('HLHRBSC08', 'Urban', 'Central 1', 90), ('HISBBSC13', 'Urban', 'North', 90), ('HISBBSC08', 'Urban', 'North', 90),
    ('HDGKBSC02', 'Urban', 'Central 3', 90), ('HPSWBSC05', 'Urban', 'North', 90), ('HMLTBSC05', 'Urban', 'Central 3', 90),
    ('HKHIBSC05', 'Urban', 'South', 90), ('HHYDBSC06', 'Urban', 'South', 90), ('HKHIBSC03', 'Urban', 'South', 90),
    ('HKHIBSC11', 'Urban', 'South', 90), ('HMLTBSC09', 'Urban', 'Central 3', 90),('HISBBSC14', 'Urban', 'North', 90), 
    ('HISBBSC15', 'Urban', 'North', 90), ('HISBBSC12', 'Urban', 'North', 90),('HKHIBSC07', 'Urban', 'South', 90),
    ('HKHIBSC14', 'Urban', 'South', 90),('HLHRBSC10', 'Urban', 'Central 1', 90),('HGJWBSC03', 'Urban', 'Central 1', 90),
    ('HKHIBSC06', 'Urban', 'South', 90),('HKHIBSC04', 'Urban', 'South', 90),('HLHRBSC12', 'Urban', 'Central 1', 90),
    ('HKHIBSC16', 'Urban', 'South', 90), ('HLHRBSC14', 'Urban', 'Central 1', 90), ('HKHIBSC12', 'Urban', 'South', 90),
    ('HFSDBSC05', 'Urban', 'Central 2', 90), ('HKHIBSC13', 'Urban', 'South', 90),
    ('HLHRBSC15', 'Urban', 'Central 1', 90), ('HLHRBSC09', 'Urban', 'Central 1', 90),
    ('HMLTBSC02','Rural', 'Central 3', 83),('HMLTBSC01','Rural', 'Central 3', 83),('HSUKBSC01','Rural', 'South', 83),
    ('HLHRBSC01','Rural', 'Central 1', 83),('HLHRBSC02','Rural', 'Central 1', 83),
    ('HISBBSC01','Rural', 'North', 83),('HISBBSC02','Rural', 'North', 83),
    ('HKHIBSC01','Rural', 'South', 83),('HKHIBSC08','Rural', 'South', 83)]

# Define column names
columns = ['GBSC', 'Type', 'Region', 'Target']

# Create a DataFrame
df1 = pd.DataFrame(g, columns=columns)
```

## Merge KPIs with BSC Defination


```python
df2 = pd.merge(df0,df1,how='left',on=['GBSC'])
```

## Re-arrange The Columns


```python
cols_to_keep = ['Date','GBSC', 'Type', 'Region', 'Target','WK','Month','Quater']
other_cols = [col for col in df2.columns if col not in cols_to_keep]
df2 = df2[cols_to_keep + other_cols]
```

## Drop Test BSC


```python
df2 = df2[~df2['GBSC'].str.contains("TEST", na=False)]
```

## Sum of Counters on Week & GBSC Group Level


```python
df3 = df2.groupby(['WK','GBSC','Type','Region','Target'])\
        [['CH331:Outgoing External Inter-Cell Handover Commands',\
          'CH333:Successful Outgoing External Inter-Cell Handovers']]\
        .sum().reset_index()
```

## Calculate Inter-BSC HSR Week Level


```python
df3['Inter-BSC HSR'] = (df3['CH333:Successful Outgoing External Inter-Cell Handovers']\
                        /df3['CH331:Outgoing External Inter-Cell Handover Commands'])*100
```

## Re-Shape Data Set


```python
df4 = df3.pivot_table\
      (index=['GBSC','Type','Region','Target'],\
      columns="WK",values='Inter-BSC HSR',aggfunc='sum').reset_index().\
     drop(['2022_52'], axis=1)
```

## Calculate Againg (Week)


```python
df4['aging_weeks']=df4.iloc[:,4:].lt(df4.iloc [:,3],axis=0).sum(axis=1)
```

## Sum of Counters on Quater & GBSC Group Level


```python
df5 = df2.groupby(['Quater','GBSC','Type','Region','Target'])\
        [['CH331:Outgoing External Inter-Cell Handover Commands',\
          'CH333:Successful Outgoing External Inter-Cell Handovers']]\
        .sum().reset_index()
```

## Calculate Inter-BSC HSR 	Quater Level


```python
df5['Inter-BSC HSR'] = (df5['CH333:Successful Outgoing External Inter-Cell Handovers']\
                        /df5['CH331:Outgoing External Inter-Cell Handover Commands'])*100
```


```python
df6 = df5.pivot_table\
      (index=['GBSC','Type','Region','Target'],\
      columns="Quater",values='Inter-BSC HSR',aggfunc='sum').reset_index()
```


```python
df6['Aging_Quater']=df6.iloc[:,4:].lt(df6.iloc [:,3],axis=0).sum(axis=1)
```

## Sum of Counters on Month & GBSC Group Level


```python
df7 = df2.groupby(['Month','GBSC','Type','Region','Target'])\
        [['CH331:Outgoing External Inter-Cell Handover Commands',\
          'CH333:Successful Outgoing External Inter-Cell Handovers']]\
        .sum().reset_index()
```

## Calculate Inter-BSC HSR 	Month Level


```python
df7['Inter-BSC HSR'] = (df7['CH333:Successful Outgoing External Inter-Cell Handovers']\
                        /df7['CH331:Outgoing External Inter-Cell Handover Commands'])*100
```


```python
df8 = df7.pivot_table\
      (index=['GBSC','Type','Region','Target'],\
      columns="Month",values='Inter-BSC HSR',aggfunc='sum').reset_index()
```


```python
df8['Aging_Month']=df8.iloc[:,4:].lt(df8.iloc [:,3],axis=0).sum(axis=1)
```

## Export Data Set


```python
with pd.ExcelWriter('Inter_BSC_HSR_KPIs.xlsx',date_format = 'dd-mm-yyyy',engine='openpyxl') as writer:

    df6.to_excel(writer,sheet_name="Inter_BSC_HSR_Quater",index=False)
    
    # Month Level KPIs Values
    df8.to_excel(writer,sheet_name="Inter_BSC_HSR_Month",index=False)
    
    # Week Level KPIs Values
    df4.to_excel(writer,sheet_name="Inter_BSC_HSR_Week",index=False)
    
    # Day (BSC Busy Hour) Level KPIs Values
    df2.to_excel(writer,sheet_name="Inter_BSC_HSR_Day(BH)",index=False)
```


```python
#re-set all the variable from the RAM
%reset -f
```

# Excel Formatting

## Import Requied Libraries


```python
# import required Libaries
import os
import openpyxl
```

## Set Input File Path


```python
path= 'D:/Advance_Data_Sets/KPIs_Analysis/Inter_BSC_HSR/KPIs'
os.chdir(path)
```

## Load Excel Sheet


```python
# Load the workbook to auto format
wb = openpyxl.load_workbook('Inter_BSC_HSR_KPIs.xlsx')
```

## Apply Filter (All the Sheets)


```python
from openpyxl.utils import get_column_letter
for ws in wb:
    # Get the first row
    first_row = ws[1]
    # Apply the filter on the first row
    ws.auto_filter.ref = f"A1:{get_column_letter(len(first_row))}1"
```

## Zoom Scale Setting (All the sheets)


```python
for ws in wb:
    ws.sheet_view.zoomScale = 95
```

## Font, Alignment and Border (All the Sheets)


```python
# Define the border style
border = openpyxl.styles.borders.Border(
    left=openpyxl.styles.borders.Side(style='thin'),
    right=openpyxl.styles.borders.Side(style='thin'),
    top=openpyxl.styles.borders.Side(style='thin'),
    bottom=openpyxl.styles.borders.Side(style='thin'))
```


```python
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle, Font, Alignment, Border, Side

# Define named styles for font, alignment, and border
style = NamedStyle(name="styled_cell")
style.font = Font(name='Calibri Light', size=8)
style.alignment = Alignment(horizontal='center', vertical='center')
style.border = Border(left=Side(style='thin'),
                      right=Side(style='thin'),
                      top=Side(style='thin'),
                      bottom=Side(style='thin'))

# Register the named style with the workbook
wb.add_named_style(style)

# Disable auto calculation
wb.calculation.calcMode = 'manual'

# Iterate through each worksheet
for ws in wb:
    # Apply named style to all cells in the sheet
    for row in ws.iter_rows():
        for cell in row:
            cell.style = "styled_cell"

# Re-enable auto calculation
wb.calculation.calcMode = 'auto'
```

## Set Date Format


```python
font_style = Font(name='Calibri Light', size=8)
alignment_style = Alignment(horizontal='center', vertical='center')

from datetime import datetime
from openpyxl.styles import NamedStyle
# Define a custom date style
date_style = NamedStyle(name='custom_date_style', number_format='DD/MM/YYYY')

# Loop through each sheet in the workbook
for ws in wb:
    # Loop through each column in the sheet
    for col in ws.iter_cols():
        for cell in col:
            if isinstance(cell.value, datetime):
                cell.style = date_style
                cell.border = border
                cell.font = font_style
                cell.alignment = alignment_style
```

## WrapText of header and Formatting (All the sheets)


```python
# wrape text of the header and center aligment
for ws in wb:
    for row in ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrapText=True, horizontal='center',vertical='center')
            cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", end_color="CF0A2C", fill_type = "solid")
            font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=11,name='Calibri Light')
            cell.font = font
```

## Set Tab Color (All the Tabs)


```python
colors = ["00B0F0", "0000FF", "ADD8E6", "87CEFA"]
for i, ws in enumerate(wb):
    ws.sheet_properties.tabColor = colors[i % len(colors)]
```

## Set Number Format (All the Sheets)


```python
for ws in wb:
    for row in ws.iter_rows(min_col=1, max_col=ws.max_column+1):
        for cell in row:
            if isinstance(cell.value,float):
                cell.number_format = '0.00'
```

## Set Column Width (All the Sheets)


```python
from openpyxl.utils import get_column_letter
# Loop through each sheet in the workbook
for ws in wb:
    # Loop through each column in the sheet
    for col in range(1, ws.max_column+1):
        # Set the width of the column to 15
        ws.column_dimensions[get_column_letter(col)].width = 15
```

## Conditional Formatting


```python
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

bold_font = openpyxl.styles.Font(bold=True,size=11,color='FFFFFF')
# empyt and none color method
empty_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')
none_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')
```


```python
ws1 = wb['Inter_BSC_HSR_Day(BH)']

last_column1 = get_column_letter(ws1.max_column-1)



for row in ws1.iter_rows(min_row=2, max_row=ws1.max_row):
    row_num = row[0].row
    formula_value = ws1.cell(row=row_num, column=5).value
    range_string = f"{last_column1}{row_num}:{last_column1}{row_num}"
    ws1.conditional_formatting.add(range_string,
                               CellIsRule(operator='lessThan',
                               formula=[formula_value],
                               fill=openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000'),
                               font=openpyxl.styles.Font(bold=True)))
    ws1.conditional_formatting.add(range_string,
                               CellIsRule(operator='greaterThanOrEqual',
                               formula=[formula_value],
                               fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00'),
                               font=openpyxl.styles.Font(bold=True)))

# Iterate through all the rows and columns
for row in ws1.iter_rows():
    for cell in row:
        if cell.value in (None, 'N/A', ' '):
            cell.fill = empty_color
        if cell.value in ('N/A', ' '):
            cell.fill = none_color
```


```python
ws2 = wb['Inter_BSC_HSR_Week']

last_column2 = get_column_letter(ws2.max_column-2)

for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row):
    row_num = row[0].row
    formula_value = ws2.cell(row=row_num, column=4).value
    range_string = f"E{row_num}:{last_column2}{row_num}"
    ws2.conditional_formatting.add(range_string,
                               CellIsRule(operator='lessThan',
                               formula=[formula_value],
                                fill=openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000'),
                               font=openpyxl.styles.Font(bold=True)))
    ws2.conditional_formatting.add(range_string,
                               CellIsRule(operator='greaterThanOrEqual',
                               #formula=[f"{formula_value}"],
                               formula=[formula_value],
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00'),
                               font=openpyxl.styles.Font(bold=True)))

# Iterate through all the rows and columns
for row in ws2.iter_rows():
    for cell in row:
        if cell.value in (None, 'N/A', ' '):
            cell.fill = empty_color
        if cell.value in ('N/A', ' '):
            cell.fill = none_color
```


```python
ws3 = wb['Inter_BSC_HSR_Month']

last_column3 = get_column_letter(ws3.max_column-2)

for row in ws3.iter_rows(min_row=2, max_row=ws3.max_row):
    row_num = row[0].row
    formula_value = ws3.cell(row=row_num, column=4).value
    range_string = f"E{row_num}:{last_column3}{row_num}"
    ws3.conditional_formatting.add(range_string,
                               CellIsRule(operator='lessThan',
                               formula=[formula_value],
                                          fill=openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000'),
                               font=openpyxl.styles.Font(bold=True)))
    ws3.conditional_formatting.add(range_string,
                               CellIsRule(operator='greaterThanOrEqual',
                               # formula=[f"{formula_value}"],
                               formula=[formula_value],
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00'),
                               font=openpyxl.styles.Font(bold=True)))

# Iterate through all the rows and columns
for row in ws3.iter_rows():
    for cell in row:
        if cell.value in (None, 'N/A', ' '):
            cell.fill = empty_color
        if cell.value in ('N/A', ' '):
            cell.fill = none_color
```


```python
ws4 = wb['Inter_BSC_HSR_Quater']

last_column4 = get_column_letter(ws4.max_column-2)

for row in ws4.iter_rows(min_row=2, max_row=ws4.max_row):
    row_num = row[0].row
    formula_value = ws4.cell(row=row_num, column=4).value
    range_string = f"E{row_num}:{last_column4}{row_num}"
    ws4.conditional_formatting.add(range_string,
                               CellIsRule(operator='lessThan',
                               formula=[formula_value],
                                          fill=openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000'),
                               font=openpyxl.styles.Font(bold=True)))
    ws4.conditional_formatting.add(range_string,
                               CellIsRule(operator='greaterThanOrEqual',
                               # formula=[f"{formula_value}"],
                               formula=[formula_value],
                                fill=openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00'),
                               font=openpyxl.styles.Font(bold=True)))

# Iterate through all the rows and columns
for row in ws4.iter_rows():
    for cell in row:
        if cell.value in (None, 'N/A', ' '):
            cell.fill = empty_color
        if cell.value in ('N/A', ' '):
            cell.fill = none_color
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
ws511.cell(row=12, column=5).value = 'Inter BSC HSR'
```

## Formatting Tital Page Report


```python
# Access the first row starting from row 3
first_row1 = list(ws511.rows)[11]
# Iterate through the cells in the first row starting from column E
for cell in first_row1[4:]:
    cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
    cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", end_color="CF0A2C", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=70,name='Calibri Light')
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

## Output


```python
# Save the changes
wb.save('Inter_BSC_HSR_KPIs.xlsx')
```

## Re-Set Variables


```python
#re-set all the variable from the RAM
%reset -f
```

## Move Final Output to Ouput Folder


```python
import shutil

# set the file path and folder paths
file_path = "D:/Advance_Data_Sets/KPIs_Analysis/Inter_BSC_HSR/KPIs/Inter_BSC_HSR_KPIs.xlsx"
destination_folder = "D:/Advance_Data_Sets/Output_Folder"

# use the shutil.move() function to move the file
shutil.move(file_path, destination_folder)
```
