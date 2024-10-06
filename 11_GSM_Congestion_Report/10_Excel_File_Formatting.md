#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# GSM RF Untilization (Excel File Formatting)

## Import Required Libraries


```python
import os
import zipfile
import openpyxl
import numpy as np
import pandas as pd
from glob import glob
```

## Set Path


```python
# Set  Path
folder_path = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'
os.chdir(folder_path)
```

## Import Excel FIle


```python
# Load the workbook to auto format
wb = openpyxl.load_workbook('GSM_RF_Utilization.xlsx')
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

# Loop through each cell in all the worksheet and set the border style
for ws in wb:
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
```

## Apply Font Style and Size (All the Sheets)


```python
# set the Font name and size of the Font of all the sheets
for ws in wb:
    for row in ws.iter_rows():
        for cell in row:
            cell.font = openpyxl.styles.Font(name='Calibri Light', size=11)
```

## Apply alignment (All the sheets)


```python
for ws in wb:
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
```

## WrapText of header (All the sheets)


```python
# wrape text of the header and center aligment
for ws in wb:
    for row in ws.iter_rows(min_row=2, max_row=2):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrapText=True, horizontal='center',vertical='center')
```

## Set Zoom Size (All the sheets)


```python
for ws in wb:
    ws.sheet_view.zoomScale = 80
```

## Set Float Number Format (All the Tabs)


```python
for ws in wb:
    for row in ws.iter_rows(min_col=1, max_col=ws.max_column):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = '0.00'
```

# Seleted Requird Sheet


```python
ws311 = wb['RF_Utilization']
```

## Delete the row


```python
# Delete the row
ws311.delete_rows(3)
```

## Hide the Column


```python
# hide the 1st column
#ws311.column_dimensions.group('A', hidden=True)
```

## Apply Filter on 2nd Row


```python
from openpyxl.utils import get_column_letter
for ws311 in wb:
    # Get the first row
    first_row = ws311[2]
    # Apply the filter on the first row
    ws311.auto_filter.ref = f"A2:{get_column_letter(len(first_row))}2"
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
ws511.cell(row=12, column=5).value = '2G RF Utilization'
```

## Formatting Tital Page Report


```python
# Access the first row starting from row 3
first_row1 = list(ws511.rows)[11]
# Iterate through the cells in the first row starting from column E
for cell in first_row1[4:]:
    cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
    cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", end_color="CF0A2C", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=80,name='Calibri Light')
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

## Clear the Cell Value


```python
# Select the sheet by index
wss = wb.worksheets[1]
# Get the cells at row 4 in the last column
cell_4 = wss.cell(row=4, column=ws2.max_column)
# Clear the contents of the cells
cell_4.value = None
# Remove the borders of the cells
cell_4.border = None
```

## Final Output


```python
# Save the changes
wb.save('GSM_RF_Utilization.xlsx')
```

## Re-Set Variables


```python
# re-set all the variable from the RAM
%reset -f
```

## Move Final Output to Ouput Folder


```python
import shutil

# set the file path and folder paths
file_path = "D:/Advance_Data_Sets/GSM_Utilization_Output_Files/GSM_RF_Utilization.xlsx"
destination_folder = "D:/Advance_Data_Sets/Output_Folder"

# use the shutil.move() function to move the file
shutil.move(file_path, destination_folder)
```


```python
# re-set all the variable from the RAM
%reset -f
```


```python
import os
import shutil

working_directory = 'D:/Advance_Data_Sets/GSM_Utilization_Output_Files'

# Delete the directory and all its contents
shutil.rmtree(working_directory, ignore_errors=True, onerror=None)
```
