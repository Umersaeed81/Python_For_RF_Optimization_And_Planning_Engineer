#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# GSM KPIs WCMS

## Import Required Libraries


```python
# Required Libraries
import os
import numpy as np
import pandas as pd
from glob import glob
from datetime import datetime, timedelta

import warnings
warnings.simplefilter("ignore")
```

## Filter the Required Files


```python
# List all files in the path
file_list = glob('D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/GSM/*.zip')

# Calculate the date for three days ago
required_date = datetime.now() - timedelta(days=7)
required_date_str = required_date.strftime("%d%m%Y")

# Filter files that match the required date or the two days before
filtered_file_list = [file for file in file_list if any((required_date + timedelta(days=i)).strftime("%d%m%Y") in file for i in range(7))]

if filtered_file_list:
    print("Filtered File List:", filtered_file_list)
else:
    print(f"No file found with the required dates in the last three days.")
```

## Import and Concat Cell DA KPIs


```python
# import & concat all the Cell DA KPIs
df=pd.concat((pd.read_csv(file,skiprows=range(6),\
    skipfooter=1,engine='python',
    parse_dates=["Date"],na_values=['NIL','/0'],\
                usecols=['Date',
                        'GBSC',
                        'Site Name',
                        'Cell Name',
                        'CSSR_Non Blocking',
                        'TCH Availability Rate(%)',
                        'Interference Band Proportion (4~5)(%)',
                        'HSR (Incoming & Outgoing)',
                        'RL9A08:Rate of Transmitted Error Frames(%)'])\
                        for file in filtered_file_list)).reset_index(drop=True)
```

## Re-Shpare KPIs (melt)


```python
# Use the melt function to transform the DataFrame
df0 = pd.melt(df, \
              id_vars=['Date','Site Name','Cell Name'], \
              value_vars=['CSSR_Non Blocking', 
                         'TCH Availability Rate(%)',
                          'Interference Band Proportion (4~5)(%)',
                          'HSR (Incoming & Outgoing)',
                          'RL9A08:Rate of Transmitted Error Frames(%)'],
              var_name='KPIs', value_name='Value').dropna()
```

## Re-Shpare KPIs (pivot_table)


```python
df1 = df0.pivot_table(index=['Site Name','Cell Name','KPIs'],columns='Date',values='Value').reset_index()
```

## Calculate KPIs


```python
df1['Average_KPIs'] = df1.iloc[:, 3:].apply(lambda row: np.average(row.dropna()), axis=1)
```

## Export Output


```python
# Set the output file path
os.chdir('D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/GSM')
# Output File
df1.to_excel('GSM_KPIs_WCMS.xlsx',index=False, sheet_name='GSM_DA_KPIs')
```


```python
# re-set all the variable from the RAM
%reset -f
```

## Load Input File 


```python
import os
import openpyxl
from openpyxl import load_workbook
os.chdir('D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/GSM')

# Load the workbook to auto format
wb = load_workbook('GSM_KPIs_WCMS.xlsx')
```

## Set Tab Color (All the Tabs)


```python
colors = ["00B0F0", "0000FF", "ADD8E6", "87CEFA"]
for i, ws in enumerate(wb):
    ws.sheet_properties.tabColor = colors[i % len(colors)]
```

## Apply border (All the Sheets)


```python
# Define the border style
border = openpyxl.styles.borders.Border(
    left=openpyxl.styles.borders.Side(style='thin'),
    right=openpyxl.styles.borders.Side(style='thin'),
    top=openpyxl.styles.borders.Side(style='thin'),
    bottom=openpyxl.styles.borders.Side(style='thin'))
```

## Font, Alignment and Border (All the Sheets)


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
border = Border(left=Side(style='thin'), \
                right=Side(style='thin'), \
                top=Side(style='thin'), \
                bottom=Side(style='thin'))

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

## Set Number Format (All the Sheets)


```python
for ws in wb:
    for row in ws.iter_rows(min_col=1, max_col=ws.max_column):
        for cell in row:
            if isinstance(cell.value,float):
                cell.number_format = '0.00'
```

## Set Zoom Size (All the sheets)


```python
for ws in wb:
    ws.sheet_view.zoomScale = 80
```

## WrapText of header and Formatting (All the sheets)


```python
# wrape text of the header and center aligment
for ws in wb:
    for row in ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrapText=True, horizontal='center',vertical='center')
            cell.fill = openpyxl.styles.PatternFill(start_color="C00000", end_color="C00000", fill_type = "solid")
            font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=11,name='Calibri Light')
            cell.font = font            
```

## Set Filter on the Header (All the sheet)


```python
from openpyxl.utils import get_column_letter
for ws in wb:
    # Get the first row
    first_row = ws[1]
    # Apply the filter on the first row
    ws.auto_filter.ref = f"A1:{get_column_letter(len(first_row))}1"
```

## Set Column Width (All the sheets)


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
ws511.merge_cells(start_row=12, start_column=5, end_row=18, end_column=17)
```

## Fill the Merge Cells


```python
# Fill the text 'LTE Cell Level KPIs Summary' in the merged cell
ws511.cell(row=12, column=5).value = 'GSM Cell DA KPIs'
```

## Formatting Tital Page Report


```python
# Access the first row starting from row 3
first_row1 = list(ws511.rows)[11]
# Iterate through the cells in the first row starting from column E
for cell in first_row1[4:]:
    cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
    cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", end_color="CF0A2C", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=60,name='Calibri Light')
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
for sheet in wb:
    if sheet.title != "Title Page":
        hyperlink_cell = ws511.cell(row=row, column=5)
        hyperlink_cell.value = sheet.title
        hyperlink_cell.hyperlink = "#'{}'!A1".format(sheet.title)
        hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
        #hyperlink_cell.border = border
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
for i, sheet in enumerate(wb.worksheets):
    # Check if the sheet is not the Title Page
    if sheet.title != "Title Page":
        # Add hyperlink to cell in the last column+2 of the sheet
        hyperlink_cell = sheet.cell(row=2, column=sheet.max_column+2)
        hyperlink_cell.value = "Back to Table of Contents"
        hyperlink_cell.hyperlink = "#'{}'!E{}".format("Title Page", 22)
        hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
        hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        
        # Set border on the hyperlink column
        for row in sheet.iter_rows(min_row=2, max_row=4, min_col=hyperlink_cell.column, max_col=hyperlink_cell.column):
            for cell in row:
                cell.border = border
        # Set width of the hyperlink column
        col_letter = openpyxl.utils.get_column_letter(hyperlink_cell.column)
        sheet.column_dimensions[col_letter].width = 25

        # Add hyperlink to cell in the last column+2 of the sheet for next sheet
        if i < len(wb.worksheets)-1:
            next_hyperlink_cell = sheet.cell(row=3, column=sheet.max_column)
            next_hyperlink_cell.value = "Next Sheet"
            next_hyperlink_cell.hyperlink = "#'{}'!A1".format(wb.worksheets[i+1].title)
            next_hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
            next_hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        
        # Add hyperlink to cell in the last column+2 of the sheet for previous sheet
        if i > 0:
            prev_hyperlink_cell = sheet.cell(row=4, column=sheet.max_column)
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

## Final Output


```python
# Save the changes
wb.save('GSM_KPIs_WCMS.xlsx')
```


```python
# re-set all the variable from the RAM
%reset -f
```


```python
import shutil

# set the file path and folder paths
file_path = "D:/Advance_Data_Sets/KPIs_Analysis/DA_KPIs_Cell_Level/GSM/GSM_KPIs_WCMS.xlsx"
destination_folder = "D:/Advance_Data_Sets/Output_Folder"

# use the shutil.move() function to move the file
shutil.move(file_path, destination_folder)
```
