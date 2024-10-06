#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# External Interference Tracker Parser

|RACaseSerialNo|Region|City|Category|2G_CI_PTML|3G_U900_CI_PTML|3G_U2100_CI_PTML|4G_CI_PTML|
|--------------|----|-------|--------|--|----|----|------|
|523/C|Center|Lahore|Jail|1234;5678|9123;1239|7865;22556|7894;55886;44456|

## Import Required Libraries & Set File Path


```python
# Python libraries
import os
import pandas as pd
```

## Set Input File Path


```python
# Set Working Folder Path
path = 'D:/Advance_Data_Sets/EI_Parser'
os.chdir(path)
```

## Import External Interference Tracker


```python
df1 = pd.read_excel('Tracker.xlsx',sheet_name=0)
```

## Re-Shapre(Melt) Data Set


```python
df2=pd.melt(df1,\
    id_vars=['RACaseSerialNo', 'Region','City', 'Category'],\
    var_name="Technology", value_name='Cell IDs')
```

## Data Pre-Processing


```python
df2.loc[:,'Cell IDs']  = df2['Cell IDs'].str.replace(",",";")
df2.loc[:,'Cell IDs'] = df2['Cell IDs'].str.replace(" ","")
df2.loc[:,'Cell IDs'] = df2['Cell IDs'].str.\
                replace(r'[^A-Za-z0-9,;]+', '', regex=True)
df2.loc[:,'Cell IDs']  = df2['Cell IDs'].str.rstrip(';')
df2.loc[:,'Cell IDs']  = df2['Cell IDs'].str.replace(";;","_")
df2.loc[:,'Cell IDs']  = df2['Cell IDs'].str.replace(";","_")
```


```python
df3 = df2[df2['Cell IDs'].notna()].reset_index(drop=True)
```

## Re-Shape (pivot_table format)


```python
df4 = (df3.pivot_table(index=['RACaseSerialNo','Region','City','Category'], 
                      columns=['Technology'], values='Cell IDs', 
                      aggfunc=lambda x: ''.join(str(v) for v in x))).reset_index().fillna('')
```

## Replace underscore with semicolon


```python
df4[['2G_CI_PTML', '3G_U2100_CI_PTML', '3G_U900_CI_PTML', '4G_CI_PTML' ]] = \
        df4[['2G_CI_PTML', '3G_U2100_CI_PTML', '3G_U900_CI_PTML', '4G_CI_PTML' ]].\
                                        applymap(lambda x: str(x).replace("_", ";"))
```

## Convert the Column to List


```python
df3 = df3.copy()
```


```python
df3.loc[:, 'Cell IDs']  = df3['Cell IDs'].str.split('_').apply(list)
```

## Re-Shapre(Explode) Data Set


```python
df5=df3.explode('Cell IDs').reset_index(level=-1, drop=True)
```

## Export Output


```python
with pd.ExcelWriter('External_Interference_Cell_List.xlsx') as writer:
    for value in df5.Technology.unique():
        df5[df5.Technology == value].\
        to_excel(writer, index=False, sheet_name=f'{value}')
    df4.to_excel(writer,sheet_name='PTML_Tracker',index=False)
```

## Exel File Formatting


```python
# import required Libarries
import openpyxl
# Load the workbook to auto format
wb = openpyxl.load_workbook('External_Interference_Cell_List.xlsx')
```

## Insert Column 


```python
ws1 = wb['2G_CI_PTML']
ws1.cell(row=1, column=ws1.max_column + 1).value = "BSC"
```


```python
ws2 = wb['3G_U900_CI_PTML']
ws2.cell(row=1, column=ws2.max_column + 1).value = "RNC"
```


```python
ws3 = wb['3G_U2100_CI_PTML']
ws3.cell(row=1, column=ws3.max_column + 1).value = "RNC"
```


```python
ws4 = wb['4G_CI_PTML']
ws4.cell(row=1, column=ws4.max_column + 1).value = "CellName"
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
            cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center',wrapText=True)
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

## Set Tab Color (All the Tabs)


```python
colors = ["00B0F0", "FFFF00", "00FF00", "FF00FF"]
for i, ws in enumerate(wb):
    ws.sheet_properties.tabColor = colors[i % len(colors)]
```

## Convert String to Number (All the Tabs)


```python
# specify the columns to be converted (column F)
column_letters = ['A','B','C','D','E','F','G','H']
# loop through each sheet in the workbook
for sheet in wb:
    # select the current sheet
    ws = sheet
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

## Column Width (All the Tabs)


```python
# Iterate over all sheets in the workbook
for ws in wb:
    for col in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].width = 25
```

## Insert a New Sheet (as First Sheet)


```python
 ws5=wb.create_sheet("Title Page", 0)
```

## Merge Specific Row and Columns


```python
ws5.merge_cells(start_row=12, start_column=5, end_row=18, end_column=17)
```

## Fill the Merge Cells


```python
# Fill the text 'CE Utilization Report' in the merged cell
ws5.cell(row=12, column=5).value = 'External Interference Tracker Parser'
```

## Formatting Tital Page Report


```python
# Access the first row starting from row 3
first_row1 = list(ws5.rows)[11]
# Iterate through the cells in the first row starting from column E
for cell in first_row1[4:]:
    cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
    cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", end_color="CF0A2C", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=45,name='Calibri Light')
    cell.font = font
```

## Inset Image


```python
from openpyxl.drawing.image import Image

# inset the Huawei logo
img = Image('D:/Advance_Data_Sets/KPIs_Analysis/Huawei.jpg')
img.width = 7 * 15
img.height = 7 * 15
ws5.add_image(img,'E3')

# inset the PTCL logo
img1 = Image('D:/Advance_Data_Sets/KPIs_Analysis/PTCL.png')
ws5.add_image(img1,'M3')
```

## Hide the gridlines


```python
ws5.sheet_view.showGridLines = False
```

## Hide the headings


```python
ws5.sheet_view.showRowColHeaders = False
```

## Hyper Link For Title Page


```python
# loop through all sheets in the workbook and insert the hyperlink to each sheet
row = 22
for sheet in wb:
    if sheet.title != "Title Page":
        hyperlink_cell = ws5.cell(row=row, column=5)
        hyperlink_cell.value = sheet.title
        hyperlink_cell.hyperlink = "#'{}'!A1".format(sheet.title)
        hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
        hyperlink_cell.border = border
        hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        # set the height of the cell
        ws5.row_dimensions[row].height = 15
        # set the colum width of column 5
        ws5.column_dimensions[get_column_letter(5)].width = 30
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
sheet = wb.worksheets[1]
# Get the cell at row 4 and the last column
cell = sheet.cell(row=4, column=sheet.max_column)
# Clear the contents of the cell
cell.value = None
# Remove the border of the cell
cell.border = None
```

## Final Output


```python
# Save the changes
wb.save('External_Interference_Cell_List.xlsx')
```

## Re-Set Variables


```python
# re-set all the variable from the RAM
%reset -f
```
