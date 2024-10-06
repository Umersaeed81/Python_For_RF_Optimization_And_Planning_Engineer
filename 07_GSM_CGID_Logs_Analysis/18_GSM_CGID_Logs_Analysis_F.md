#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# Format Excel File

## Import required Libarries


```python
# import required Libarries
import os
import openpyxl
```

## Set Input File Path


```python
#set the Path (Path must be same format)
path = 'D:/Advance_Data_Sets/BSS_EI_Output/CDIG_Processing'
os.chdir(path)
```

## Laad Excel File


```python
# Load the workbook to auto format
wb = openpyxl.load_workbook('CDIG_Analysis.xlsx')
```

## Remove the Unwanted Sheet


```python
# Remove the sheet
wb.remove(wb['Sheet'])
```

## Set Tab Color (All the Tabs)


```python
colors = ["00B0F0", "0000FF", "ADD8E6", "87CEFA"]
for i, ws in enumerate(wb):
    ws.sheet_properties.tabColor = colors[i % len(colors)]
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

## Set Zoom Size (All the sheets)


```python
for ws in wb:
    ws.sheet_view.zoomScale = 80
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

## Set Float Number Format (All the Tabs)


```python
for ws in wb:
    for row in ws.iter_rows(min_col=1, max_col=ws.max_column):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = '0.00'
```

## Convert String to Number


```python
ws1 = wb['CDIG_Cell_Level']
```


```python
column_letters = ['B']
# select the specified sheet
#ws = wb[sheet_name]
# loop through each specified column
for column_letter in column_letters:
    # get the range of cells in the column
    column_range = ws1[column_letter]
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

## Conditional Formatting


```python
# libraries
from openpyxl.styles import Font
from openpyxl.formatting.rule import CellIsRule
# bold defination
bold_font = openpyxl.styles.Font(bold=True,size=11,name='Calibri Light')
# empyt and none color method
empty_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')
none_color = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF')
```


```python
for col in range(5, ws1.max_column):
    column_letter = get_column_letter(col)
    range_string = f"{column_letter}2:{column_letter}{ws1.max_row}"
    ws1.conditional_formatting.add(range_string,
                               CellIsRule(operator='greaterThanOrEqual',
                               formula=[15.0],
                               fill=openpyxl.styles.PatternFill(start_color='FFCCCC', end_color='FFCCCC')))
    for row in ws1[range_string]:
        for cell in row:
            if cell.value and cell.value >= 15.00:
                cell.font = bold_font

# Iterate through all the rows and columns
for row in ws1.iter_rows():
    for cell in row:
        if cell.value in (None, 'N/A', ' ',):
            cell.fill = empty_color
        if cell.value in ('N/A', ' '):
            cell.fill = none_color
```


```python
ws2 = wb['CDIG_BSC_Level']
```


```python
for col in range(2, ws2.max_column):
    column_letter = get_column_letter(col)
    range_string = f"{column_letter}2:{column_letter}{ws2.max_row}"
    ws2.conditional_formatting.add(range_string,
                               CellIsRule(operator='greaterThanOrEqual',
                               formula=[1000],
                               fill=openpyxl.styles.PatternFill(start_color='FFCCCC', end_color='FFCCCC')))
    for row in ws2[range_string]:
        for cell in row:
            if cell.value and cell.value >= 1000:
                cell.font = bold_font

# Iterate through all the rows and columns
for row in ws2.iter_rows():
    for cell in row:
        if cell.value in (None, 'N/A', ' ',):
            cell.fill = empty_color
        if cell.value in ('N/A', ' '):
            cell.fill = none_color
```


```python
ws3 = wb['GSM_BTS_Ava']
```


```python
from openpyxl.formatting.rule import CellIsRule
for col in range(3, ws3.max_column):
    column_letter = get_column_letter(col)
    range_string = f"{column_letter}2:{column_letter}{ws3.max_row}"
    ws3.conditional_formatting.add(range_string,
                               CellIsRule(operator='lessThan',
                               formula=[100],
                               fill=openpyxl.styles.PatternFill(start_color='FFCCCC', end_color='FFCCCC')))
    for row in ws3[range_string]:
        for cell in row:
            if cell.value and cell.value < 100:
                cell.font = bold_font

# Iterate through all the rows and columns
for row in ws3.iter_rows():
    for cell in row:
        if cell.value in (None, 'N/A', ' ',):
            cell.fill = empty_color
        if cell.value in ('N/A', ' '):
            cell.fill = none_color
```


```python
ws4 = wb['GSM_Ava_24hrs_Interval_Count']
```


```python
for col in range(3, ws4.max_column):
    column_letter = get_column_letter(col)
    range_string = f"{column_letter}2:{column_letter}{ws4.max_row}"
    ws4.conditional_formatting.add(range_string,
                               CellIsRule(operator='greaterThanOrEqual',
                               formula=[5.0],
                               fill=openpyxl.styles.PatternFill(start_color='FFCCCC', end_color='FFCCCC')))
    for row in ws4[range_string]:
        for cell in row:
            if cell.value and cell.value >= 5.00:
                cell.font = bold_font

# Iterate through all the rows and columns
for row in ws4.iter_rows():
    for cell in row:
        if cell.value in (None, 'N/A', ' ',):
            cell.fill = empty_color
        if cell.value in ('N/A', ' '):
            cell.fill = none_color
```


```python
ws5 = wb['GSM_Ava_9-21hrs_Interval_Count']
```


```python
for col in range(3, ws5.max_column):
    column_letter = get_column_letter(col)
    range_string = f"{column_letter}2:{column_letter}{ws5.max_row}"
    ws5.conditional_formatting.add(range_string,
                               CellIsRule(operator='greaterThanOrEqual',
                               formula=[5.0],
                               fill=openpyxl.styles.PatternFill(start_color='FFCCCC', end_color='FFCCCC')))
    for row in ws5[range_string]:
        for cell in row:
            if cell.value and cell.value >= 5.00:
                cell.font = bold_font

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
# Fill the text 'CE Utilization Report' in the merged cell
ws511.cell(row=12, column=5).value = 'GSM CDIG Log Analysis'
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

## Hyper Link For Title Page


```python
# Define the border style
border = openpyxl.styles.borders.Border(
    left=openpyxl.styles.borders.Side(style='thin'),
    right=openpyxl.styles.borders.Side(style='thin'),
    top=openpyxl.styles.borders.Side(style='thin'),
    bottom=openpyxl.styles.borders.Side(style='thin'))

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

## Hide the gridlines


```python
ws511.sheet_view.showGridLines = False
```

## Hide the headings


```python
ws511.sheet_view.showRowColHeaders = False
```

## Final Output


```python
# Save the changes
wb.save('CDIG_Analysis.xlsx')
```

## Re-Set Variables


```python
# re-set all the variable from the RAM
%reset -f
```


```python
import shutil

# set the file path and folder paths
file_path = "D:/Advance_Data_Sets/BSS_EI_Output/CDIG_Processing/CDIG_Analysis.xlsx"
destination_folder = "D:/Advance_Data_Sets/Output_Folder"

# use the shutil.move() function to move the file
shutil.move(file_path, destination_folder)
```

## Delete Files From the Folder, If Folder Not Found Created the Folder


```python
import os
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
