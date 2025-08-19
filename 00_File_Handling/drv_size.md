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
base_path = "E:/"
```

## Get list of top-level folders


```python
folders = [f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f))]

data = []
for folder in folders:
    folder_path = os.path.join(base_path, folder)
    size_bytes = get_folder_size(folder_path)
    size_gb = size_bytes / (1024**3)  # Convert to GB
    data.append([folder, round(size_gb, 2), date.today()])
```

## Create DataFrame


```python
df = pd.DataFrame(data, columns=["Folder", "Size_GB", "Date"])
```

## Export Output


```python
folder_path = 'D:/Otput'
os.chdir(folder_path)

today_str = date.today().strftime("%Y%m%d")
df.to_excel(f"File_size_{today_str}.xlsx",index=False)
```
