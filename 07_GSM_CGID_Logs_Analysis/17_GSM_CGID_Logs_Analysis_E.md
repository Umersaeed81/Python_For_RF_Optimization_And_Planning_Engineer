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
import datetime
from openpyxl.styles import NamedStyle
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import warnings
warnings.simplefilter("ignore")
```

## User Define Funcation to Combaine All the Excel Sheet


```python
def copy_sheet(source_sheet, target_sheet):
    for row in source_sheet.iter_rows(min_row=1, max_row=source_sheet.max_row, min_col=1, max_col=source_sheet.max_column, values_only=True):
        new_row = []
        for cell in row:
            if isinstance(cell, (int, float)):
                new_row.append(cell)
            elif isinstance(cell, str):
                new_row.append(cell)
            elif isinstance(cell, (datetime.date, datetime.datetime)):
                formatted_date = cell.strftime('%d-%m-%Y')
                new_row.append(formatted_date)
            else:
                new_row.append(None)
        target_sheet.append(new_row)

def copy_workbook(folder_path, new_file_path):
    new_workbook = Workbook()

    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            workbook = load_workbook(file_path, data_only=True)

            for sheet_name in workbook.sheetnames:
                source_sheet = workbook[sheet_name]

                # Skip blank sheets
                if source_sheet.max_row == 0 or source_sheet.max_column == 0:
                    continue

                new_sheet = new_workbook.create_sheet(title=sheet_name)

                # Copy values and styles
                copy_sheet(source_sheet, new_sheet)

                # Copy merged cells
                for merged_cells_range in source_sheet.merged_cells.ranges:
                    min_col, min_row, max_col, max_row = merged_cells_range.min_col, merged_cells_range.min_row, merged_cells_range.max_col, merged_cells_range.max_row
                    new_sheet.merge_cells(start_row=min_row, end_row=max_row, start_column=min_col, end_column=max_col)

    new_workbook.save(new_file_path)
```

## Apply on the Required Folder


```python
folder_path = 'D:/Advance_Data_Sets/BSS_EI_Output/CDIG_Processing'
new_file_path = 'D:/Advance_Data_Sets/BSS_EI_Output/CDIG_Processing/CDIG_Analysis.xlsx'
copy_workbook(folder_path, new_file_path)
```


```python
#re-set all the variable from the RAM
%reset -f
```
