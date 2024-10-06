#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# LTE NodB Level Licence

## Format Excel File


```python
import os
folder_path = 'D:/Advance_Data_Sets/License/LTE_NodeB_Level'
os.chdir(folder_path)
```

## Load Excel File


```python
# import required Libaries
import openpyxl
# Load the workbook to auto format
wb = openpyxl.load_workbook('LTE_NodeB_Level_Licence.xlsx')
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

## Convert String to Number (All the Tabs)


```python
# Loop through each sheet in the workbook
for ws in wb:
    # Loop through each column in the sheet
    for col in ws.iter_cols():
        for cell in col:
            # Check if the cell value is a string
            if isinstance(cell.value, str):
                try:
                    # Convert the string to a number
                    cell.value = float(cell.value)
                except ValueError:
                    # If the string cannot be converted to a float, ignore it
                    pass
```

## Set Column Width (All the Sheets)


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
# Fill the text 'LTE NodeB Level Licence' in the merged cell
ws511.cell(row=12, column=5).value = 'LTE NodeB Level Licence'
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

## Export Output


```python
# Save the changes
wb.save('LTE_NodeB_Level_Licence.xlsx')
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
file_path = "D:/Advance_Data_Sets/License/LTE_NodeB_Level/LTE_NodeB_Level_Licence.xlsx"
destination_folder = "D:/Advance_Data_Sets/Output_Folder"

# use the shutil.move() function to move the file
shutil.move(file_path, destination_folder)
```
