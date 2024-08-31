#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>

# User Address Extraction in WCMS Complaints 
![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/00_WCMS_Tille.png?raw=true)

## Task Description

PTML has provided the WCMS complaints data in an Excel file for analysis. The objective of this Python script is to extract the 'User Address' information from the 'Standard Comments' column.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/01_WCMS_Task_Description.png?raw=true)

## Import Required Libraries

The code imports necessary libraries and modules for data manipulation and Excel file formatting. Specifically:<br>
**1. `os`:** For handling operating system-related tasks like changing directories.<br>
**2. `openpyxl and load_workbook`:** To work with Excel files, including loading, reading, and modifying them.<br>
**3. `pandas`:** For data analysis and manipulation, especially when dealing with data frames.<br>
**4. `openpyxl.utils and openpyxl.styles`:** To facilitate formatting tasks in Excel, such as adjusting column widths, setting fonts, aligning text, and applying borders.<br>
**5. `Suppress Warnings`:** The warnings filter is set to ignore, ensuring that any warning messages do not clutter the output.<br>
This setup creates a robust environment for reading, processing, and formatting Excel data files efficiently.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/02_WCMS_Libraries.png?raw=true)


```python
import os
import openpyxl
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle, Font, Alignment, Border, Side
import warnings
warnings.simplefilter("ignore")
```

## Set Input File Path

The code sets the input path to `'D:/Advance_Data_Sets/WCMS'` and changes the current working directory to this specified path. This ensures that subsequent file operations are performed in the correct directory.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/03_WCMS_Input_File_Path.png?raw=true)


```python
# Set the Input Path 
path = 'D:/Advance_Data_Sets/WCMS'
os.chdir(path)
```

## Import Input File

The code reads data from an Excel file named `'WCMS.xlsx'` into a pandas DataFrame. It specifies that the second row of the file should be used as the header and selects only the columns `'MSISDN'`, `'Complaint Id'`, `'Region'`, and `'Standard Comments'` for inclusion in the DataFrame.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/04_WCMS_Import_DataSet.png?raw=true)


```python
df = pd.read_excel('WCMS.xlsx',\
                   header=1,\
                   usecols=['MSISDN','Complaint Id','Region','Standard Comments'])
```

## Code for Extracting and Filtering Location Information from Complaint Data

The code performs the following steps:

**1. Convert 'Standard Comments' Column to List:** Splits the entries in the `'Standard Comments'` column by the '|' character and converts them into lists.

**2. Filter 'Party A' Elements:** Filters out the comments that contain 'Party A' from each list and creates a new column `'Filtered Comments'` with these filtered comments.


**3. Extract Address Information:** Extracts the location from the filtered comments. It splits each comment by the ';' character and retrieves the second part (if available), storing it in a new column `'Location'`.

**4. Filter Required Columns:** Selects only the relevant columns (`'MSISDN'`, `'Complaint Id'`, `'Region'`, and `'Location'`) and stores the result in a new DataFrame `df1`.


**5. Export Output:** Saves the filtered DataFrame `df1` to a CSV file named `'Location.csv'`, excluding the index column.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/05_WCMS_Data_Processing.png?raw=true)


```python
# Convert 'Standard Comments' Column to list
df.loc[:, 'Standard Comments']  = df['Standard Comments'].str.split('|').apply(list)
```


```python
# Filter Party A Element From the list
df['Filtered Comments'] = df['Standard Comments'].\
                          apply(lambda comments: [comment for comment in comments if 'Party A' in comment])
```


```python
# Extract only the Address
df['Address'] = df['Filtered Comments'].apply(lambda comments: comments[0].split(';')[1] if comments else None)
```


```python
# Filter Required Columns
df1 = df[['MSISDN','Complaint Id','Region','Address']]
```


```python
# Convert the 'MSISDN' column to string to prevent scientific notation
df1['MSISDN'] = df1['MSISDN'].astype(str)
```


```python
# Export Output File
df1.to_excel('WCMS_Address.xlsx',index=False,sheet_name='WCMS_Address')
```

## Excel File Formatting

The code loads an existing Excel workbook named `'WCMS_Address.xlsx'` using the openpyxl library. This allows for reading, modifying, and applying formatting changes to the workbook in subsequent steps.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/06_Load_Excel_Using_Openpyxl.png?raw=true)


```python
# Load the workbook to auto format
wb = openpyxl.load_workbook('WCMS_Address.xlsx')
```

## Set Tab Color (All the Tabs)

The code sets the tab color of each worksheet in the loaded Excel workbook ('`wb'`) to one of the specified colors from the `colors` list. It iterates through all worksheets in the workbook, and for each worksheet, assigns a color from the list ["`00B0F0`", "`0000FF`", "`ADD8E6`", "`87CEFA`"]. The colors are applied cyclically using the modulo operator (`%`), ensuring that the color pattern repeats if there are more worksheets than colors.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/07_Tab_Colors.png?raw=true)


```python
colors = ["00B0F0", "0000FF", "ADD8E6", "87CEFA"]
for i, ws in enumerate(wb):
    ws.sheet_properties.tabColor = colors[i % len(colors)]
```

## Font, Alignment and Border (All the Sheets)

The code customizes the formatting of all cells in each worksheet of the loaded Excel workbook (`wb`) using the `openpyxl` library. The steps are as follows:

**1. Define the Border Style:** Creates a thin border style for all sides (left, right, top, and bottom) of a cell.

**2. Define Named Styles:** A named style called `styled_cell` is created with:

   - **Font**: 'Calibri Light' with a size of 8.
   - **Alignment**: Centered both horizontally and vertically.
   - **Border**: A thin border on all sides, as defined earlier.

**3. Register the Named Style:** The named style (`styled_cell`) is added to the workbook (`wb`) so it can be applied throughout.

**4. Disable Auto Calculation:** Temporarily disables Excel's auto-calculation mode by setting `calcMode` to `'manual'` to improve performance while formatting cells.

**5. Apply the Named Style to All Cells:** Iterates through each worksheet and applies the `styled_cell` style to every cell in the sheet.

**6. Re-enable Auto Calculation:** Restores the default auto-calculation mode after formatting is complete by setting `calcMode` to `'auto'`.

This code ensures that all cells in the workbook are formatted consistently with the specified font, alignment, and border style while optimizing the performance during the formatting process.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/08_Excel_Cells_Formatting.png?raw=true)


```python
# Define the border style
border = openpyxl.styles.borders.Border(
    left=openpyxl.styles.borders.Side(style='thin'),
    right=openpyxl.styles.borders.Side(style='thin'),
    top=openpyxl.styles.borders.Side(style='thin'),
    bottom=openpyxl.styles.borders.Side(style='thin'))

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

The code formats the header row (the first row) of each worksheet in the loaded Excel workbook (`wb`) using the `openpyxl` library. It performs the following formatting tasks:

**1. Wrap Text and Center Alignment:** For each header cell, the text is set to wrap within the cell boundaries, and the text alignment is set to be centered both horizontally and vertically.

**2. Apply Background Fill Color:** The background color of each header cell is set to a solid red color (C00000), providing a visual distinction for the headers.

**3. Set Font Style:** The font style of each header cell is customized to be white (`FFFFFF`) in color, bold, with a size of 11, and using the 'Calibri Light' font type.

By applying these styles, the code ensures that the header row in every worksheet is visually distinct, well-formatted, and easy to read.


![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/09_Formatting_Header_Cells.png?raw=true)




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

The code applies an auto-filter to the first row (header row) of each worksheet in the loaded Excel workbook (`wb`) using the `openpyxl` library. It performs the following actions:

**1. Get the First Row:** Retrieves the first row (`first_row`) of each worksheet (`ws`), which is typically used as the header row.

**2. Apply Auto-Filter:** Sets an auto-filter on the first row, covering all columns from A1 to the last column of the first row. The `get_column_letter()` function is used to dynamically determine the last column letter based on the length of `first_row`.

This code adds filtering capabilities to the header row in every worksheet, allowing users to sort and filter the data easily when viewing the Excel file.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/10_Auto_filter.png?raw=true)


```python
for ws in wb:
    # Get the first row
    first_row = ws[1]
    # Apply the filter on the first row
    ws.auto_filter.ref = f"A1:{get_column_letter(len(first_row))}1"
```

## Set Column Width (All the sheets)

The code adjusts the width of each column in all worksheets of the loaded Excel workbook (`wb`) to fit the content more appropriately. It performs the following steps:

**1. Iterate Over All Sheets:** Loops through each worksheet in the workbook.

**2. Iterate Over All Columns:** Within each worksheet, iterates through all columns.

**3. Determine Column Width:**

   - **Current Width:** Retrieves the current width of the column.
   - **Maximum Cell Width:** Calculates the maximum length of the content in each cell of the column, converting each cell’s value to a string to determine its length.

**4. Adjust Column Width:** If the maximum cell width exceeds the current column width, updates the column width to match the maximum content length.

This ensures that each column is resized to accommodate its longest cell content, improving the readability and presentation of the data in the Excel file.

![](
https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/11_column_width.png?raw=true)


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

The code customizes a newly created worksheet titled "Title Page" in the loaded Excel workbook (`wb`) with specific formatting and content:

**1. Create New Sheet:** Adds a new worksheet named "Title Page" at the beginning of the workbook.

**2. Merge Cells:** Merges cells from row 12, column 5 to row 18, column 17, creating a larger cell area.

**3. Fill Merged Cell:** Inserts the text 'WCMS Complaints' into the merged cell at row 12, column 5.

**4. Format Header Row:**

   - **Select Header Row:** Accesses the row starting from row 3 (the 12th row of the worksheet).
   - **Apply Formatting:** Iterates through the cells starting from column E (the 5th column), setting their alignment to center, applying a red background color (`CF0A2C`), and formatting the font to white, bold, size 60, and using 'Calibri Light'.

**5. Insert Logos:**

   - **Huawei Logo:** Inserts an image of the Huawei logo from the specified path (`'D:/Advance_Data_Sets/KPIs_Analysis/Huawei.jpg'`) and sizes it appropriately.<br>
   - **PTCL Logo:** Inserts an image of the PTCL logo from the specified path (`'D:/Advance_Data_Sets/KPIs_Analysis/PTCL.png'`).

**6. Hide Grid Lines and Headers:** Configures the worksheet to hide grid lines and row/column headers for a cleaner appearance.

This code sets up a well-formatted title page with specific visual elements and customization for the workbook.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/12_T_page_formatting.png?raw=true)


```python
ws511 =wb.create_sheet("Title Page", 0)
```

## Merge Specific Row and Columns


```python
ws511.merge_cells(start_row=12, start_column=5, end_row=18, end_column=17)
```

## Fill the Merge Cells


```python
# Fill the text 'WCMS Complaints' in the merged cell
ws511.cell(row=12, column=5).value = 'WCMS Complaints'
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

The code adds hyperlinks to navigate between sheets in the Excel workbook (wb) and enhances user navigation features:

**1. Add Hyperlinks to Title Page:**

  - **Insert Hyperlinks:** Loops through all sheets in the workbook (excluding the "Title Page") and adds a hyperlink to each sheet name in the "Title Page" worksheet at row 22. The hyperlink directs users to the respective sheet.
  - **Format Hyperlinks:** Sets the hyperlink text color to blue and underlined, centers the text within the cell, adjusts the row height to 15, and sets the column width of column 5 to 30.

**2. Add Navigation Links in Each Sheet:**

  - **Back to Table of Contents:** Adds a hyperlink in each sheet's last column plus two, directing users back to the "Title Page" at cell E22. The hyperlink text is styled in blue with an underline and centered.

  - **Previous and Next Sheet Navigation:** Adds links in each sheet to navigate to the previous and next sheets if applicable:
    - **Next Sheet:** A hyperlink directs users to the next sheet if it exists.
    - **Previous Sheet:** A hyperlink directs users to the previous sheet if it exists.
    - **Format Navigation Links:** Sets a border around the hyperlink cells and adjusts the column width to 25.

This setup enhances the workbook by providing easy navigation between different sheets and returning to the main "Title Page," improving usability and accessibility.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/13_hyperlinks.png?raw=true)


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

## Final Output

The code saves all changes made to the workbook (`wb`) back to the file named `'WCMS_Address.xlsx'`.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/15_save.png?raw=true)


```python
# Save the changes
wb.save('WCMS_Address.xlsx')
```

## Re-set all the variable

The code `%reset -f` is an IPython magic command used to reset the Python environment by clearing all variables, functions, imports, and other user-defined objects from memory (RAM). The `-f` flag stands for "force" and ensures that this action is performed without asking for user confirmation.

By using `%reset -f`, the code effectively provides a clean slate in the current IPython session or Jupyter Notebook, which is useful for preventing potential conflicts, freeing up memory, and avoiding any unintended use of previous variables or objects.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Diagram/WCMS_Complaints/14_reset_variables.png?raw=true)


```python
# re-set all the variable from the RAM
%reset -f
```
