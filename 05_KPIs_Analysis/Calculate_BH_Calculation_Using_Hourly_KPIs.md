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
import os
import glob
import pandas as pd

import warnings
warnings.simplefilter("ignore")
```

## Input File Path


```python
# Get the list of CSV and XLSX files
csv_files = glob.glob('D:/BH_Calculation/InputFiles/*.csv')
xlsx_files = glob.glob('D:/BH_Calculation/InputFiles/*.xlsx')
```

## Variable Reset based on CSV and Excel File Presence


```python
# Get the list of CSV and XLSX files
csv_files = glob.glob('D:/BH_Calculation/InputFiles/*.csv')
xlsx_files = glob.glob('D:/BH_Calculation/InputFiles/*.xlsx')

if csv_files and xlsx_files:
    print ('Both .csv and .xlsx found, further working is not allowed.')
    %reset -f
```

## Functions for Processing CSV and Excel Files with Date and Time Parsing


```python
def process_csv_with_date(files):
    return pd.concat((pd.read_csv(file, skiprows=range(6), skipfooter=1, engine='python', 
                                 parse_dates=["Date"], na_values=['NIL','/0']) for file in files)).\
                                 sort_values('Date').reset_index(drop=True)

def process_csv_with_time(files):
    return pd.concat((pd.read_csv(file, skiprows=range(6), skipfooter=1, engine='python', 
                                  parse_dates=["Time"], na_values=['NIL','/0']) for file in files)).\
                                  sort_values('Time').reset_index(drop=True)

def process_xlsx_with_date(files):
    return pd.concat((pd.read_excel(file, engine='openpyxl',
                                    converters={'Integrity': lambda value: '{:,.0f}%'.format(value * 100)},
                                    parse_dates=["Date"], na_values=['NIL','/0']) for file in files)).\
                                    sort_values('Date').reset_index(drop=True)

def process_xlsx_with_time(files):
    return pd.concat((pd.read_excel(file, engine='openpyxl', 
                                    converters={'Integrity': lambda value: '{:,.0f}%'.format(value * 100)},
                                    parse_dates=["Time"], na_values=['NIL','/0']) for file in files)).\
                                    sort_values('Time').reset_index(drop=True)
```

## File Format Identification for Date and Time Columns in CSV and Excel Files


```python
def check_file_formats(csv_files, xlsx_files):
    date_csv_files = []
    time_csv_files = []
    date_xlsx_files = []
    time_xlsx_files = []
    
    for file in csv_files:
        with open(file) as f:
            lines = f.readlines()
            if "Date" in lines[6]:
                date_csv_files.append(file)
            elif "Time" in lines[6]:
                time_csv_files.append(file)
    
    for file in xlsx_files:
        df = pd.read_excel(file, engine='openpyxl')
        if "Date" in df.columns:
            date_xlsx_files.append(file)
        elif "Time" in df.columns:
            time_xlsx_files.append(file)
    
    return date_csv_files, time_csv_files, date_xlsx_files, time_xlsx_files
```

## Data Import and Concatenation Based on File Formats


```python
def import_and_concatenate(csv_files, xlsx_files):
    date_csv_files, time_csv_files, date_xlsx_files, time_xlsx_files = check_file_formats(csv_files, xlsx_files)
    
    if (date_csv_files or date_xlsx_files) and (time_csv_files or time_xlsx_files):
        print("Both csv and xlsx formats of files found. Not importing and concatenating data.")
        return None
    elif date_csv_files or date_xlsx_files:
        print("Files with date format found. Importing and concatenating data with date column.")
        if date_csv_files:
            date_data = process_csv_with_date(date_csv_files)
        else:
            date_data = process_xlsx_with_date(date_xlsx_files)
        return date_data
    elif time_csv_files or time_xlsx_files:
        print("Files with time format found. Importing and concatenating data with time column.")
        if time_csv_files:
            time_data = process_csv_with_time(time_csv_files)
        else:
            time_data = process_xlsx_with_time(time_xlsx_files)
        return time_data
    else:
        print("No suitable files found.")
```

## Data Import and Concatenation with Error Handling


```python
df0 = pd.DataFrame()  # Initialize an empty DataFrame

if csv_files or xlsx_files:
    try:
        # Attempt to import and concatenate the data
        df0 = import_and_concatenate(csv_files, xlsx_files)
    except Exception as e:
        print("Error: Multiple file formats not supported by the System.")
```

## Calculate BH KPIs


```python
# df0 is your input DataFrame
if 'Date' in df0.columns:
    # Case-1: DataFrame has 'Date' column
    # Calculate CS Traffic cluster BH
    df1 = df0.loc[df0.groupby(['Date', 'UCell Group'])['VS.RAB.AMR.Erlang.cell(Erl)'].idxmax()].reset_index(drop=True)
    # Calculate PS Traffic cluster BH
    df2 = df0.loc[df0.groupby(['Date', 'UCell Group'])['PS Traffic(GB)'].idxmax()].reset_index(drop=True)
elif 'Time' in df0.columns:
    # Case-2: DataFrame has 'Time' column
    df0['Date'] = df0['Time'].dt.date  # Extract date from Time column
    # Calculate CS Traffic cluster BH
    df1 = df0.loc[df0.groupby(['Date', 'UCell Group'])['VS.RAB.AMR.Erlang.cell(Erl)'].idxmax()].reset_index(drop=True)
    # Calculate PS Traffic cluster BH
    df2 = df0.loc[df0.groupby(['Date', 'UCell Group'])['PS Traffic(GB)'].idxmax()].reset_index(drop=True)
else:
    raise ValueError("The DataFrame must contain either 'Date' or 'Time' column.")
```

## Export Each Day Busy Hour


```python
# Set Export File Path
folder_path = 'D:/BH_Calculation/OutputFile'
os.chdir(folder_path)

# Check if files list is empty
files = os.listdir()
if not files:
    new_max_num = 0
else:
    # Extract the maximum number from existing files
    max_num = max(int(file.split('_')[0]) for file in files)
    # Increment the maximum number
    new_max_num = max_num + 1

# Iterate over unique dates and export to separate Excel files with multiple sheets
for unique_date in df1['Date'].unique():
    unique_date_str = unique_date.strftime('%d%m%Y')
    
    # Generate the output Excel file name
    output_file = f'{new_max_num:02d}_UMTS_Cluster_BH_{unique_date_str}.xlsx'
    
    # Filter the DataFrames for the current unique date
    df1_filtered = df1[df1['Date'] == unique_date]
    df2_filtered = df2[df2['Date'] == unique_date]

    # Check if the DataFrame contains a datetime column
    if any(df1_filtered.dtypes == 'datetime64[ns]'):
        date_format = 'dd-mm-yyyy HH:MM'
        datetime_format = 'dd-mm-yyyy HH:MM'
    else:
        date_format = 'dd-mm-yyyy'
        datetime_format = 'dd-mm-yyyy'

    # Create Excel writer object
    with pd.ExcelWriter(output_file, date_format=date_format, datetime_format=datetime_format) as writer:
        # Write DataFrame to Excel with sheet_name "UMTS_CS_BH"
        df1_filtered.to_excel(writer, sheet_name="UMTS_CS_BH", index=False)
        # Write DataFrame to the same Excel file with sheet_name "UMTS_PS_BH"
        df2_filtered.to_excel(writer, sheet_name="UMTS_PS_BH", index=False)

    # Increment the maximum number for the next iteration
    new_max_num += 1
```

## Merge All Busy Hour Files in Single File


```python
# List ALl the Excel Files
all_files = glob.glob('D:/BH_Calculation/OutputFile/*_UMTS_Cluster_BH_*.xlsx')

# Get the Sheet Name
sheets = pd.ExcelFile(all_files[0]).sheet_names

# concat and import all Excel Files
dfs = {s: pd.concat((pd.read_excel(f, sheet_name=s ) for f in all_files), \
                    ignore_index=True) for s in sheets} 
```

## Export Merge Cluster Busy Hour File


```python
# Set Export File Path
folder_path = 'D:/BH_Calculation'
os.chdir(folder_path)

with pd.ExcelWriter('UMTS_Cluster_BH.xlsx') as writer:
    for keys in dfs:
        dfs[keys].to_excel(writer,keys,index=False)     
```

## Excel File Formatting


```python
# import required Libaries
import openpyxl
# Load the workbook to auto format
wb = openpyxl.load_workbook('UMTS_Cluster_BH.xlsx')
#-------------------------------------------------------------------------------------#     
# Apply Filter (All the Sheets)
from openpyxl.utils import get_column_letter
for ws in wb:
    # Get the first row
    first_row = ws[1]
    # Apply the filter on the first row
    ws.auto_filter.ref = f"A1:{get_column_letter(len(first_row))}1"
#-------------------------------------------------------------------------------------#  
# Zoom Scale Setting (All the sheets)
for ws in wb:
    ws.sheet_view.zoomScale = 95
#-------------------------------------------------------------------------------------#              
from openpyxl.styles import NamedStyle
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from datetime import datetime, time, date

# Define custom date styles
date_style = NamedStyle(name='custom_date_style', number_format='dd-mm-yyyy')
datetime_style = NamedStyle(name='custom_datetime_style', number_format='dd-mm-yyyy HH:MM')


# Loop through each sheet in the workbook
for ws in wb:
    # Identify the header row
    header_row = ws[1]

    # Loop through each column in the sheet
    for col_idx, col in enumerate(ws.iter_cols()):
        # Check the style of the header cell in the corresponding column
        header_style = header_row[col_idx].value

        # Check if any cell in the column contains date or time
        contains_date = any(isinstance(cell.value, date) for cell in col)
        contains_time = any(isinstance(cell.value, time) for cell in col)

        if header_style == 'Date':
            # Apply date style if no time values are present, otherwise apply datetime style
            if not contains_time:
                for cell in col:
                    if isinstance(cell.value, date):
                        cell.style = date_style
            else:
                for cell in col:
                    if isinstance(cell.value, (date, datetime)):
                        cell.style = datetime_style
#-------------------------------------------------------------------------------------#  
# Apply border (All the Sheets)
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
#-------------------------------------------------------------------------------------# 
# Apply Font Style and Size (All the Sheets)
# set the Font name and size of the Font of all the sheets
for ws in wb:
    for row in ws.iter_rows():
        for cell in row:
            cell.font = openpyxl.styles.Font(name='Calibri Light', size=11)
#-------------------------------------------------------------------------------------# 
# Apply alignment (All the sheets)
for ws in wb:
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
#-------------------------------------------------------------------------------------# 
# WrapText of header and Formatting (All the sheets)
# wrape text of the header and center aligment
for ws in wb:
    for row in ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrapText=True, horizontal='center',vertical='center')
            cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", end_color="CF0A2C", fill_type = "solid")
            font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=11,name='Calibri Light')
            cell.font = font
#-------------------------------------------------------------------------------------#
# Set Tab Color (All the Tabs)
colors = ["00B0F0", "0000FF", "ADD8E6", "87CEFA"]
for i, ws in enumerate(wb):
    ws.sheet_properties.tabColor = colors[i % len(colors)]
#-------------------------------------------------------------------------------------#
# Set Number Format (All the Sheets)
for ws in wb:
    for row in ws.iter_rows(min_col=1, max_col=ws.max_column+1):
        for cell in row:
            if isinstance(cell.value,float):
                cell.number_format = '0.00'
#-------------------------------------------------------------------------------------#
#Set Column Width (All the sheets)
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
#-------------------------------------------------------------------------------------#            
# Insert a New Sheet (as First Sheet)
ws511 =wb.create_sheet("Title Page", 0)
#-------------------------------------------------------------------------------------#
# Merge Specific Row and Columns
ws511.merge_cells(start_row=12, start_column=5, end_row=18, end_column=17)
#-------------------------------------------------------------------------------------#
# Fill the Merge Cells
# Fill the text 'Cluster Busy Hour Report' in the merged cell
ws511.cell(row=12, column=5).value = 'Cluster Busy Hour Report'
#-------------------------------------------------------------------------------------#
# Formatting Tital Page Report
# Access the first row starting from row 3
first_row1 = list(ws511.rows)[11]
# Iterate through the cells in the first row starting from column E
for cell in first_row1[4:]:
    cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
    cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", end_color="CF0A2C", fill_type = "solid")
    font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=50,name='Calibri Light')
    cell.font = font
#-------------------------------------------------------------------------------------#
# Inset Image
from openpyxl.drawing.image import Image

# inset the Huawei logo
img = Image('D:/BH_Calculation/logo/Huawei.jpg')
img.width = 7 * 15
img.height = 7 * 15
ws511.add_image(img,'E3')

# inset the PTCL logo
img1 = Image('D:/BH_Calculation/logo/jazz.png')
ws511.add_image(img1,'M3')
#-------------------------------------------------------------------------------------#
# Hide the gridlines
ws511.sheet_view.showGridLines = False
#-------------------------------------------------------------------------------------#
# Hide the headings
ws511.sheet_view.showRowColHeaders = False
#-------------------------------------------------------------------------------------#
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
#-------------------------------------------------------------------------------------#
# Hyper Link For Sub Pages
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
#-------------------------------------------------------------------------------------#
# Clear the Cell Value
# Select the sheet by index
wss = wb.worksheets[1]
# Get the cell at row 4 and the last column
cell = wss.cell(row=4, column=wss.max_column)
# Clear the contents of the cell
cell.value = None
# Remove the border of the cell
cell.border = None
#-------------------------------------------------------------------------------------#
# Output
# Save the changes
wb.save('UMTS_Cluster_BH.xlsx')
#-------------------------------------------------------------------------------------#
#re-set all the variable from the RAM
%reset -f
#-------------------------------------------------------------------------------------#
```
