## Import the requied liraries


```python
import os
import pandas as pd
from datetime import date
```

## User define funcation to calculate size of the folder


```python
def get_folder_size(path):
    """Return total size of folder in bytes."""
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(path):
        for f in filenames:
            try:
                fp = os.path.join(dirpath, f)
                total_size += os.path.getsize(fp)
            except Exception:
                pass
    return total_size
```

## Set your Base Path


```python
base_path = "D:/Advance_Data_Sets"
```

## Calculate Folder Size


```python
# --- Sheet 1: Base folder sizes ---
folders = [f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f))]

data_base = []
data_child = []

for folder in folders:
    folder_path = os.path.join(base_path, folder)
    
    # Base folder size
    size_bytes = get_folder_size(folder_path)
    size_gb = size_bytes / (1024**3)
    data_base.append([folder, round(size_gb, 2), date.today()])

    # --- Sheet 2: Child folder sizes ---
    child_folders = [cf for cf in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, cf))]
    for child in child_folders:
        child_path = os.path.join(folder_path, child)
        child_size_bytes = get_folder_size(child_path)
        child_size_gb = child_size_bytes / (1024**3)
        data_child.append([folder, child, round(child_size_gb, 2), date.today()])
```

## DataFrames


```python
df_base = pd.DataFrame(data_base, columns=["Base_Folder", "Size_GB", "Date"])
df_child = pd.DataFrame(data_child, columns=["Base_Folder", "Child_Folder", "Size_GB", "Date"])
```

## Export Oput


```python
output_path = 'D:/Otput'
os.makedirs(output_path, exist_ok=True)

today_str = date.today().strftime("%Y%m%d")
file_path = os.path.join(output_path, f"File_size_{today_str}.xlsx")

with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
    df_base.to_excel(writer, sheet_name="Base_Folders", index=False)
    df_child.to_excel(writer, sheet_name="Child_Folders", index=False)

print("✅ Excel file created:", file_path)
```

    ✅ Excel file created: D:/Otput\File_size_20250821.xlsx
    
