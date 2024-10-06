#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

## Import Required Libraries


```python
import openpyxl
from glob import glob
```

## Load Required File


```python
# Use glob to find the file with the pattern
file_list = glob('D:/Advance_Data_Sets/GUL/GUL_Output/10_UMTS_LTE_Capacity_Analysis_*.xlsx')

# Check if any files matched the pattern
if file_list:
    # Load the first matched file
    wb = openpyxl.load_workbook(file_list[0], read_only=False)
    print("Workbook loaded:", file_list[0])
else:
    print("No files found with the pattern '10_UMTS_LTE_Capacity_Analysis_*.xlsx'")
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

## Header Formatting


```python
# Select Specific Sheet
ws4000 = wb['Site_Mapping']

# Get the maximum number of columns with data in the first row
max_column = ws4000.max_column

# Iterate over the cells in the first row up to the last column with data
for row in ws4000.iter_rows(min_row=1, max_row=1, max_col=max_column-4):
    for cell in row:
        cell.alignment = openpyxl.styles.Alignment(wrapText=True, horizontal='center', vertical='center')
        cell.fill = openpyxl.styles.PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        font = openpyxl.styles.Font(color="FFFFFF", bold=True, size=11, name='Calibri Light')
        cell.font = font
```

## Set Column Width


```python
# Iterate over all columns in the 'Site_Mapping' sheet
for column in ws4000.columns:
    # Get the current width of the column
    current_width = ws4000.column_dimensions[column[0].column_letter].width
    
    # Get the maximum width of the cells in the column
    length = max(len(str(cell.value)) for cell in column) + 4
    
    # Set the width of the column to fit the maximum width, if it's greater than the current width
    if length > current_width:
        ws4000.column_dimensions[column[0].column_letter].width = length
```

## Select Specific Sheet


```python
ws311 = wb['GUL Utilization']
```

## Column Width of Merge Cells Sheet


```python
from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter


# Iterate over all columns in the sheet
for column in ws311.columns:
    # Check if the column contains any merged cells
    if ws311.merged_cells.ranges:
        for merged_range in ws311.merged_cells.ranges:
            start_column = get_column_letter(merged_range.min_col)
            end_column = get_column_letter(merged_range.max_col)
            if start_column <= get_column_letter(column[0].column) <= end_column:
                # Get the current width of the column
                current_width = ws311.column_dimensions[get_column_letter(column[0].column)].width
                # Get the maximum width of the cells in the column, skipping the header row
                length = max(len(str(cell.value)) for cell in column[1:])
                # Set the width of the column to fit the maximum width, if it's greater than the current width
                if length > current_width:
                    ws311.column_dimensions[get_column_letter(column[0].column)].width = length
```

## Delete Specific Row


```python
# Delete the row
ws311.delete_rows(3)
```

## Delete Column in Multi-index Header DataFrame


```python
def delete_col_with_merged_ranges(ws311, idx):
    ws311.delete_cols(idx)
    for mcr in ws311.merged_cells:
        if idx < mcr.min_col:
            mcr.shift(col_shift=-1)
        elif idx <= mcr.max_col:
            mcr.shrink(right=1)
# Delete column 'A' by name
delete_col_with_merged_ranges(ws311, 1)  # Specify the column index (1 for column 'A')
```

## Header Formatting


```python
# Get the maximum number of columns with data in the first row
max_column = ws311.max_column


# wrape text of the header and center aligment
for ws311 in wb:
    for row in ws.iter_rows(min_row=1, max_row=2,max_col=max_column-1):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrapText=True, horizontal='center',vertical='center')
            cell.fill = openpyxl.styles.PatternFill(start_color="C00000", end_color="C00000", fill_type = "solid")
            font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=11,name='Calibri Light')
            cell.font = font
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
# Fill the text 'UMTS LTE Capacity Analysis' in the merged cell
ws511.cell(row=12, column=5).value = 'UMTS LTE Capacity Analysis'
```

## Formatting Tital Page Report


```python
# Access the first row starting from row 3
first_row1 = list(ws511.rows)[11]
# Iterate through the cells in the first row starting from column E
for cell in first_row1[4:]:
    cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
    cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", end_color="CF0A2C", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=50,name='Calibri Light')
    cell.font = font
```

## Inset Image


```python
from openpyxl.drawing.image import Image

# inset the Huawei logo
img = Image('D:/MAK/logo/Huawei.jpg')
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
# Define the border style
border = openpyxl.styles.borders.Border(
    left=openpyxl.styles.borders.Side(style='thin'),
    right=openpyxl.styles.borders.Side(style='thin'),
    top=openpyxl.styles.borders.Side(style='thin'),
    bottom=openpyxl.styles.borders.Side(style='thin'))

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

## Select Active Sheet


```python
for sheet in wb.worksheets:
    sheet.views.sheetView[0].tabSelected = False  # Deselect all sheets
wb.active = wb['Title Page']  # Set the title page as the active sheet
```

## Save Output


```python
import os
from datetime import datetime
# Set output File path
path = 'D:/Advance_Data_Sets/GUL/GUL_Output'
os.chdir(path)

# Get today's date in the format 'ddmmyyyy'
today_date = datetime.today().strftime('%d%m%Y')

# Create the filename with today's date
filename = f'10_UMTS_LTE_Capacity_Analysis_{today_date}.xlsx'

# Save the workbook with the new filename
wb.save(filename)
```

## Re-set All Variables


```python
# re-set all the variable from the RAM
%reset -f
```

## Move Output File to the Output Folder


```python
import shutil
from glob import glob

# Find all files matching the pattern
file_paths = glob('D:/Advance_Data_Sets/GUL/GUL_Output/10_UMTS_LTE_Capacity_Analysis_*.xlsx')

# Set the destination folder
destination_folder = "D:/Advance_Data_Sets/Output_Folder"

# Move each file individually
for file_path in file_paths:
    shutil.move(file_path, destination_folder)
```
