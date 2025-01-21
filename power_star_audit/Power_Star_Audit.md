#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

-----------------------------------------------

# GSM Power Star Audit

## Import Required Libraries


```python
# Python libraries
import os
import zipfile
import numpy as np
import pandas as pd
from glob import glob
import warnings
warnings.simplefilter("ignore")
```

## Delete Files From the Folder, If Folder Not Found Created the Folder


```python
# Folder path
folder_path = r"D:/Advance_Data_Sets/RF_Export/GSM/Output"

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

    Deleted file: D:/Advance_Data_Sets/RF_Export/GSM/Output\Power_Star_Audit.xlsx
    

## Unzip GSM RF Export (Frequency Export)


```python
folder_path = 'D:/Advance_Data_Sets/RF_Export/GSM/00_Frequency_Export'
os.chdir(folder_path)

# Python code to unzip the Folders
for file in os.listdir(folder_path):   # get the list of files
    if zipfile.is_zipfile(file): # if it is a zipfile, extract it
        with zipfile.ZipFile(file) as zip_ref: # treat the file as a zip
            folder_name = os.path.splitext(file)[0]  # get the name of the folder (remove '.zip' extension)
            os.makedirs(folder_name, exist_ok=True)  # create the folder if it doesn't exist
            zip_ref.extractall(folder_name)  # extract the contents into the folder
```

## RF Export (GTRX)


```python
# get Cell File Path from all the Sub folders
df = glob('D:/Advance_Data_Sets/RF_Export/GSM/00_Frequency_Export/**/GTRX.txt', recursive=True)

# import and concat GTRX File
df0=pd.concat((pd.read_csv(file,header=1,\
                 engine='python',encoding='unicode_escape',\
                 usecols=['BSC Name', 'Cell Name','Frequency', 'Is Main BCCH TRX','Active Status','TRX ID'],\
                 dtype={'Cell Name':str,'TRX ID':str}) for file in df)).reset_index(drop=True)


# Band Identification
df0['Band'] = np.where(
            ((df0['Frequency']>=25) & (df0['Frequency']<=62)),
            'PTML-GSM-TRX', 
            np.where(
                    (df0['Frequency']>=512) & (df0['Frequency']<=586), 
                    'PTML-DCS-TRX', 
                     'Other-Band-TRX'))

# Filter BCCH TRXs
df1 =df0[df0['Is Main BCCH TRX']=='YES'].reset_index(drop=True)
```

## Unzip GSM RF Export (Cell Export)


```python
# Set Working Folder Path For 2G RF Export (Cell Level)
path= 'D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export'
os.chdir(path)

# Python code to unzip the Folders
for file in os.listdir(path):   # get the list of files
    if zipfile.is_zipfile(file): # if it is a zipfile, extract it
        with zipfile.ZipFile(file) as zip_ref: # treat the file as a zip
            folder_name = os.path.splitext(file)[0]  # get the name of the folder (remove '.zip' extension)
            os.makedirs(folder_name, exist_ok=True)  # create the folder if it doesn't exist
            zip_ref.extractall(folder_name)  # extract the contents into the folder
```

## RF Export (Cell-File)


```python
# get Cell File Path from all the Sub folders
df2 = glob('D:/Advance_Data_Sets/RF_Export/GSM/00_Frequency_Export/**/GCELL.txt', recursive=True)

# import and concat Cell File
df3=pd.concat((pd.read_csv(file,header=1,\
                 engine='python',encoding='unicode_escape',\
                 usecols=['BSC Name','Cell CI','Cell Name','active status','BTS Name'],\
                 dtype={'Cell Name':str}) for file in df2)).reset_index(drop=True)

# Final Cell File
df4 = pd.merge(df3,df1[['BSC Name','Cell Name','Band']],how='left',on=['BSC Name','Cell Name']).fillna('-')
```

## 2G-TRX Power Amp-Cell & 2G-TRX Voltage Adjust-BSC


```python
# get Cell File Path from all the Sub folders
df5 = glob('D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export/**/GCELLBASICPARA.txt', recursive=True)
# import and concat GCELLBASICPARA File
df6=pd.concat((pd.read_csv(file,header=1,\
                 engine='python',encoding='unicode_escape',\
                 usecols=['BSC Name','Cell Name','Cell Index','Allow Dynamic Shutdown of TRX','Adjust Voltage'],\
                 dtype={'Cell Index':str,'Cell Name':str}) for file in df5)).reset_index(drop=True).\
                rename(columns = {'Allow Dynamic Shutdown of TRX':'DYNOPENTRXPOWER','Adjust Voltage':'BTSADJUST'})
# merge GCELLBASICPARA with GCELL
df7 = pd.merge(df4,df6,how='left',on=['BSC Name','Cell Name']).fillna('-')
```

## 2G-Enhanced BCCH Pwr-Cell


```python
# get Cell File Path from all the Sub folders
df8 = glob('D:/Advance_Data_Sets/RF_Export/GSM/04_Cell_Export/**/GCELLOTHEXT.txt', recursive=True)

#import and concat GCELLOTHEXT File
df9=pd.concat((pd.read_csv(file,header=1,\
                 engine='python',encoding='unicode_escape',\
                 usecols=['BSC Name','Cell Name','Cell Index','Power Derating Enabled'],\
                 dtype={'Cell Index':str,'Cell Name':str}) for file in df8)).reset_index(drop=True).\
                rename(columns = {'Power Derating Enabled':'MAINBCCHPWRDTEN'})

# merge GCELLOTHEXT with GCELL
df10= pd.merge(df4,df9,how='left',on=['BSC Name','Cell Name']).fillna('-')
```

## Unzip GSM RF Export (GTRX Export)


```python
folder_path = 'D:/Advance_Data_Sets/RF_Export/GSM/03_GTRX_Export'
os.chdir(folder_path)

# Python code to unzip the Folders
for file in os.listdir(folder_path):   # get the list of files
    if zipfile.is_zipfile(file): # if it is a zipfile, extract it
        with zipfile.ZipFile(file) as zip_ref: # treat the file as a zip
            folder_name = os.path.splitext(file)[0]  # get the name of the folder (remove '.zip' extension)
            os.makedirs(folder_name, exist_ok=True)  # create the folder if it doesn't exist
            zip_ref.extractall(folder_name)  # extract the contents into the folder
```

## 2G-TRX Power Amp-TRX 


```python
# get Cell File Path from all the Sub folders
df11 = glob('D:/Advance_Data_Sets/RF_Export/GSM/03_GTRX_Export/**/GTRXDEV.txt', recursive=True)
#import and concat GTRXDEV File
df12=pd.concat((pd.read_csv(file,header=1,\
                 engine='python',encoding='unicode_escape',\
                 usecols=['BSC Name','BTS Name','Cell Name','Cell Index', 'Is Main BCCH TRX',
                'Cell CI','TRX ID','Allow Dynamic Shutdown of TRX'],\
                 dtype={'Cell Index':str,'Cell Name':str,'TRX ID':str}) for file in df11)).reset_index(drop=True).\
                rename(columns = {'Allow Dynamic Shutdown of TRX':'CPS'})

# merge with GTRX
df13 = pd.merge(df12,df0[['BSC Name','Cell Name','TRX ID','Active Status','Band','Frequency']],\
                how='left',on=['BSC Name','Cell Name','TRX ID']).\
                rename(columns={'Active Status': 'Active Status(TRX Level)','Band':'TRX_Band'})

# merge with GCELL
df14 = pd.merge(df13,df4[['BSC Name', 'BTS Name', 'Cell Name','active status']],\
                 how='left',on=['BSC Name', 'BTS Name', 'Cell Name']).\
                 rename(columns={'active status': 'Active Status(Cell Level)'})   
```

## 2G-TRX Power Amp-BSC


```python
df15 = glob('D:/Advance_Data_Sets/RF_Export/GSM/05_Config_Data/ConfigurationData_*.xlsx',recursive=True)
df16 = (
    pd.read_excel(
        f,
        sheet_name='BSCDSTPA',
        header=[1],
        engine='openpyxl',
        usecols=['*BSC Name', 'Allow Dynamic Shutdown of TRX by BSC']
    )
    for f in df15
)
df17 = pd.concat(df16, ignore_index=True)
```

## Export Output


```python
# Set Output File Path
path = 'D:/Advance_Data_Sets/RF_Export/GSM/Output'
os.chdir(path)
# Export Output
with pd.ExcelWriter('Power_Star_Audit_GSM.xlsx') as writer:
    # 2G-TRX Power Amp-BSC
    df17.to_excel(writer,sheet_name='2G-TRX Power Amp-BSC',index=False)
    # 2G-TRX Power Amp-Cell
    df7.to_excel(writer,sheet_name='2G-TRX Power Amp-Cell',\
                 index=False,\
                 columns=['BSC Name', 'BTS Name', 
                          'Cell Name', 'Cell CI', 
                          'active status', 'Band',
                       'Cell Index', 'DYNOPENTRXPOWER'])
    # 2G-TRX Voltage Adjust-BSC
    df7.to_excel(writer,sheet_name='2G-TRX Voltage Adjust-BSC',\
                 index=False,\
                 columns=['BSC Name', 'BTS Name', 'Cell Name', 'Cell CI', 'active status', 'Band',
                  'Cell Index', 'BTSADJUST'])
    # 2G-Enhanced BCCH Pwr-Cell
    df10.to_excel(writer,sheet_name='2G-Enhanced BCCH Pwr-Cell',index=False)
    # 2G-TRX Power Amp-TRX
    df14.to_excel(writer,sheet_name='2G-TRX Power Amp-TRX',index=False)
```


```python
#re-set all the variable from the RAM
%reset -f
```

# LTE Power Star Audit

## Import Required Libraries


```python
# Python libraries
import os
import zipfile
import numpy as np
import pandas as pd
from glob import glob
import warnings
warnings.simplefilter("ignore")
```


```python
df = glob('D:/Advance_Data_Sets/RF_Export/LTE/Cell/ConfigurationData_*.xlsx',recursive=True)
```

## LTE - Symbol Power Saving_1


```python
df0 = (pd.read_excel(f, sheet_name='ENODEBALGOSWITCH',\
        header=[1], engine='openpyxl',\
       usecols=['*eNodeB Name',
       'Power save switch',
      ])\
       .assign(**{
       'Power save switch': lambda x: x['Power save switch'].str.split('&').apply(list),
       'Power save switch req': lambda x: x['Power save switch'].apply(lambda y: ', '.join([elem for elem in y if elem.startswith('SymbolShutdownSwitch-')])),
         }) for f in df)
df1 = pd.concat(df0, ignore_index=True)

# drop the column
df1.drop(columns=['Power save switch'], inplace=True)

# rename the column
df1.rename(columns={'Power save switch req': 'Power save switch'}, inplace=True)
```

## LTE - Symbol Power Saving_2


```python
df2 = (pd.read_excel(f, sheet_name='CELLALGOSWITCH',\
        header=[1], engine='openpyxl',dtype={'*Local cell ID':str},\
       usecols=['*eNodeB Name', '*Local cell ID',
       'DL schedule switch',
      ])\
       .assign(**{
       'DL schedule switch': lambda x: x['DL schedule switch'].str.split('&').apply(list),
       'DL schedule switch req': lambda x: x['DL schedule switch'].apply(lambda y: ', '.join([elem for elem in y if elem.startswith('MBSFNShutDownSwitch-')])),
         }) for f in df)
df3 = pd.concat(df2, ignore_index=True)

# drop the column
df3.drop(columns=['DL schedule switch'], inplace=True)
# rename the column
df3.rename(columns={'DL schedule switch req': 'DL schedule switch'}, inplace=True)

# Cell File Info
df4 = (
    pd.read_excel(
        f,
        sheet_name='CELL',
        header=[1],
        engine='openpyxl',
        dtype={'*Local Cell ID': str},
        usecols=['*eNodeB Name', '*Local Cell ID', '*Cell Name','Downlink EARFCN']
    )
    for f in df
)
df5 = pd.concat(df4, ignore_index=True).rename(columns={'*Local Cell ID': '*Local cell ID'})

# Band Identification
df5['Band'] = np.where(
            (df5['Downlink EARFCN'].eq(1276)),
            'PTML-L-1800', 
            np.where(
                    (df5['Downlink EARFCN'].eq(3628)), 
                    'PTML-L-900', 
                np.where(
                    (df5['Downlink EARFCN'].eq(175)), 
                    'PTML-L-2100', 
                     'Other-Band')))

# merge CELLALGOSWITCH with CELL
df6 = pd.merge(df3,df5,how='left',on=['*eNodeB Name', '*Local cell ID'])
# re-order the columns
df6 = df6[['*eNodeB Name','*Cell Name','*Local cell ID','Downlink EARFCN','Band','DL schedule switch']]
```

## Export Output


```python
# Set Output File Path
path = 'D:/Advance_Data_Sets/RF_Export/GSM/Output'
os.chdir(path)
# Export Output
with pd.ExcelWriter('Power_Star_Audit_LTE.xlsx') as writer:
    # LTE - Symbol Power Saving_1
    df1.to_excel(writer,sheet_name='LTE-Symbol Power Saving_1',index=False)
    # LTE - Symbol Power Saving_2
    df6.to_excel(writer,sheet_name='LTE-Symbol Power Saving_2',index=False)
```


```python
#re-set all the variable from the RAM
%reset -f
```

# Power Star Audit

## Import Required Libraries


```python
# Python libraries
import os
import zipfile
import numpy as np
import pandas as pd
from glob import glob
import warnings
warnings.simplefilter("ignore")
```

## Set File Path


```python
# Set File Path
path = 'D:/Advance_Data_Sets/RF_Export/GSM/Output'
os.chdir(path)
```

## GSM Power Star Audit


```python
df = pd.read_excel('Power_Star_Audit_LTE.xlsx',sheet_name='LTE-Symbol Power Saving_1')
df0 = pd.read_excel('Power_Star_Audit_LTE.xlsx',sheet_name='LTE-Symbol Power Saving_2')
```

## LTE Power Star Audit


```python
df1 = pd.read_excel('Power_Star_Audit_GSM.xlsx',sheet_name='2G-TRX Power Amp-BSC')
df2 = pd.read_excel('Power_Star_Audit_GSM.xlsx',sheet_name='2G-TRX Power Amp-Cell')
df3 = pd.read_excel('Power_Star_Audit_GSM.xlsx',sheet_name='2G-TRX Voltage Adjust-BSC')
df4 = pd.read_excel('Power_Star_Audit_GSM.xlsx',sheet_name='2G-Enhanced BCCH Pwr-Cell')
df5 = pd.read_excel('Power_Star_Audit_GSM.xlsx',sheet_name='2G-TRX Power Amp-TRX')
```

## Export 


```python
# Export Output
with pd.ExcelWriter('Power_Star_Audit.xlsx') as writer:
    # LTE Audit
    df.to_excel(writer,sheet_name='LTE-Symbol Power Saving_1',index=False)
    df0.to_excel(writer,sheet_name='LTE-Symbol Power Saving_2',index=False)
    # GSM Audit
    df1.to_excel(writer,sheet_name='2G-TRX Power Amp-BSC',index=False)
    df2.to_excel(writer,sheet_name='2G-TRX Power Amp-Cell',index=False)
    df3.to_excel(writer,sheet_name='2G-TRX Voltage Adjust-BSC',index=False)
    df4.to_excel(writer,sheet_name='2G-Enhanced BCCH Pwr-Cell',index=False)
    df5.to_excel(writer,sheet_name='2G-TRX Power Amp-TRX',index=False)
```


```python
#re-set all the variable from the RAM
%reset -f
```

# Delete Unwanted Files


```python
import os

directory_path = "D:/Advance_Data_Sets/RF_Export/GSM/Output"

files_to_delete = ["Power_Star_Audit_GSM.xlsx", 
                  "Power_Star_Audit_LTE.xlsx"]

for file_name in files_to_delete:
    file_path = os.path.join(directory_path, file_name)

    try:
        os.remove(file_path)
        print(f"File {file_name} deleted successfully.")
    except FileNotFoundError:
        print(f"File {file_name} not found.")
    except Exception as e:
        print(f"Error deleting {file_name}: {e}")
```

    File Power_Star_Audit_GSM.xlsx deleted successfully.
    File Power_Star_Audit_LTE.xlsx deleted successfully.
    

# Excel File Formatting

## Import Required Libraries


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

## Set the Input Path 


```python
# Set the Input Path 
path = 'D:/Advance_Data_Sets/RF_Export/GSM/Output'
os.chdir(path)
```

## Load Excel File


```python
# Load the workbook to auto format
wb = openpyxl.load_workbook('Power_Star_Audit.xlsx')
```

## Set Tab Color (All the Tabs)


```python
colors = ["00B0F0", "0000FF", "ADD8E6", "87CEFA"]
for i, ws in enumerate(wb):
    ws.sheet_properties.tabColor = colors[i % len(colors)]
```

## Font, Alignment and Border (All the Sheets)


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
# Fill the text 'Power Start' in the merged cell
ws511.cell(row=12, column=5).value = 'Power Star Audit'
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

## Final Output


```python
# Save the changes
wb.save('Power_Star_Audit.xlsx')
```
