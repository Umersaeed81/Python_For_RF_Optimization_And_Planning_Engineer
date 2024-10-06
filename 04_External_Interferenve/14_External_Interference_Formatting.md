#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

## Set input File Path


```python
import os
# File Path
path = 'D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files_EI'
os.chdir(path)
```

## Load Excel File


```python
# import required Libarries
import openpyxl
# Load the workbook to auto format
wb = openpyxl.load_workbook('External_Interference_KPIs.xlsx')
```

## Reassign the sheets in the workbook according to the desired order


```python
# Define the desired sheet order
desired_order = ['GSM_IOI_Huawei', 'Band1', 'Band8','LTE_UL_Interference_Huawei',
                 'GSM_IOI_Intervals','UMTS_U900_RTWP_Intervals','UMTS_U2100_RTWP_Intervals','LTE_UL_Interference_Intervals']

# Reassign the sheets in the workbook according to the desired order
wb._sheets = [wb[sheet] for sheet in desired_order]
```

## Rename the Tab Name


```python
# Define a dictionary that maps old sheet names to new sheet names
sheet_names = {"Band1": "U2100_Band1_Huawei", "Band8": "U900_Band8_Huawei"}

# Loop over the dictionary items
for old_name, new_name in sheet_names.items():
    # Select the sheet
    sheet = wb[old_name]
    # Change the name of the sheet
    sheet.title = new_name
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

## WrapText of header and Formatting (All the sheets)


```python
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font

# Assuming 'wb' is your workbook and it has been loaded already
for ws in wb:
    # Get the maximum number of columns with data
    max_column = ws.max_column-1
    for row in ws.iter_rows(min_row=1, max_row=1, max_col=max_column):
        for cell in row:
            cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
            cell.font = Font(color="FFFFFF", bold=True, size=11)
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

## Convert String to Number (All the Tabs)


```python
# specify the columns to be converted (column F)
column_letters = ['A','B','C','D','E','F','G','H']
# loop through each sheet in the workbook
for ws in wb:
#     # select the current sheet
#     ws = sheet
    # loop through each specified column
    for column_letter in column_letters:
        # get the range of cells in the column
        column_range = ws[column_letter]
        # loop through each cell in the column
        for cell in column_range:
            # check if the cell value is a string
            if type(cell.value) == str:
                try:
                    # convert the string to a number
                    cell.value = float(cell.value)
                except ValueError:
                    # if the string cannot be converted to a float, ignore it
                    pass
```

## Condition Formatting

## 1. GSM


```python
# select required sheet
ws2 = wb['GSM_IOI_Huawei']
```


```python
from itertools import chain
from openpyxl.styles import Font
from openpyxl.formatting.rule import CellIsRule
```


```python
bold_font = openpyxl.styles.Font(bold=True,size=10)
```


```python
for col in chain(range(4, 5), range(6, ws2.max_column-6)):
    column_letter = get_column_letter(col)
    range_string = f"{column_letter}2:{column_letter}{ws2.max_row}"
    ws2.conditional_formatting.add(range_string,
                               CellIsRule(operator='greaterThanOrEqual',
                               formula=[10.0],
                               fill=openpyxl.styles.PatternFill(start_color='FFCCCC', end_color='FFCCCC')))
    for row in ws2[range_string]:
        for cell in row:
            if cell.value and cell.value >= 10.00:
                cell.font = bold_font

empty_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')
none_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')
# Iterate through all the rows and columns
for row in ws2.iter_rows():
    for cell in row:
        if cell.value in (None, 'N/A', ' ',):
            cell.fill = empty_color
        if cell.value in ('N/A', ' '):
            cell.fill = none_color
```

## U-2100


```python
ws3 = wb['U2100_Band1_Huawei']
```


```python
for col in chain(range(6, 7), range(9, ws3.max_column-7)):
    column_letter = get_column_letter(col)
    range_string = f"{column_letter}2:{column_letter}{ws3.max_row}"
    ws3.conditional_formatting.add(range_string,
                               CellIsRule(operator='greaterThanOrEqual',
                               formula=[-98.0],
                               fill=openpyxl.styles.PatternFill(start_color='FFCCCC', end_color='FFCCCC')))
    for row in ws3[range_string]:
        for cell in row:
            if cell.value and cell.value >= -98.00:
                cell.font = bold_font

empty_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')
none_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')
# Iterate through all the rows and columns
for row in ws3.iter_rows():
    for cell in row:
        if cell.value in (None, 'N/A', ' ',):
            cell.fill = empty_color
        if cell.value in ('N/A', ' '):
            cell.fill = none_color
```

## U-900


```python
ws4 = wb['U900_Band8_Huawei']
```


```python
for col in chain(range(6, 7), range(9, ws4.max_column-7)):
    column_letter = get_column_letter(col)
    range_string = f"{column_letter}2:{column_letter}{ws4.max_row}"
    ws4.conditional_formatting.add(range_string,
                               CellIsRule(operator='greaterThanOrEqual',
                               formula=[-95.0],
                               fill=openpyxl.styles.PatternFill(start_color='FFCCCC', end_color='FFCCCC')))
    for row in ws4[range_string]:
        for cell in row:
            if cell.value and cell.value >= -95.00:
                cell.font = bold_font

empty_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')
none_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')
# Iterate through all the rows and columns
for row in ws4.iter_rows():
    for cell in row:
        if cell.value in (None, 'N/A', ' ',):
            cell.fill = empty_color
        if cell.value in ('N/A', ' '):
            cell.fill = none_color
```

## LTE


```python
ws5 = wb['LTE_UL_Interference_Huawei']
```


```python
for col in chain(range(3, 4), range(5, ws5.max_column-6)):
    column_letter = get_column_letter(col)
    range_string = f"{column_letter}2:{column_letter}{ws5.max_row}"
    ws5.conditional_formatting.add(range_string,
                               CellIsRule(operator='greaterThanOrEqual',
                               formula=[-108.0],
                               fill=openpyxl.styles.PatternFill(start_color='FFCCCC', end_color='FFCCCC')))
    for row in ws5[range_string]:
        for cell in row:
            if cell.value and cell.value >= -108.00:
                cell.font = bold_font

empty_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')
none_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')
# Iterate through all the rows and columns
for row in ws5.iter_rows():
    for cell in row:
        if cell.value in (None, 'N/A', ' ',):
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
# Fill the text 'External Interference KPIs' in the merged cell
ws511.cell(row=12, column=5).value = 'External Interference KPIs'
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
for ws2 in wb:
    if ws2.title != "Title Page":
        hyperlink_cell = ws511.cell(row=row, column=5)
        hyperlink_cell.value = ws2.title
        hyperlink_cell.hyperlink = "#'{}'!A1".format(ws2.title)
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
for i, ws2 in enumerate(wb.worksheets):
    # Check if the sheet is not the Title Page
    if ws2.title != "Title Page":
        # Add hyperlink to cell in the last column+2 of the sheet
        hyperlink_cell = ws2.cell(row=2, column=ws2.max_column+2)
        hyperlink_cell.value = "Back to Table of Contents"
        hyperlink_cell.hyperlink = "#'{}'!E{}".format("Title Page", 22)
        hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
        hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        
        # Set border on the hyperlink column
        for row in ws2.iter_rows(min_row=2, max_row=4, min_col=hyperlink_cell.column, max_col=hyperlink_cell.column):
            for cell in row:
                cell.border = border
        # Set width of the hyperlink column
        col_letter = openpyxl.utils.get_column_letter(hyperlink_cell.column)
        ws2.column_dimensions[col_letter].width = 25

        # Add hyperlink to cell in the last column+2 of the sheet for next sheet
        if i < len(wb.worksheets)-1:
            next_hyperlink_cell = ws2.cell(row=3, column=ws2.max_column)
            next_hyperlink_cell.value = "Next Sheet"
            next_hyperlink_cell.hyperlink = "#'{}'!A1".format(wb.worksheets[i+1].title)
            next_hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
            next_hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        
        # Add hyperlink to cell in the last column+2 of the sheet for previous sheet
        if i > 0:
            prev_hyperlink_cell = ws2.cell(row=4, column=ws2.max_column)
            prev_hyperlink_cell.value = "Previous Sheet"
            prev_hyperlink_cell.hyperlink = "#'{}'!A1".format(wb.worksheets[i-1].title)
            prev_hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
            prev_hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
```

## Final Output


```python
# Save the changes
wb.save('External_Interference_KPIs.xlsx')
```


```python
# re-set all the variable from the RAM
%reset -f
```


```python
import shutil

# set the file path and folder paths
file_path = "D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files_EI/External_Interference_KPIs.xlsx"
destination_folder = "D:/Advance_Data_Sets/Output_Folder"

# use the shutil.move() function to move the
shutil.move(file_path, destination_folder)
```
