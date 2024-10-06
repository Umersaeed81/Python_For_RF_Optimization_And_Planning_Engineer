```python
import os

directory_path = "D:/Advance_Data_Sets/BSS_EI_Output/Processing_Files"
files_to_delete = ["02_UMTS_BSS_Issus.xlsx", 
                   "03_UMTS_BSS_Issus_Intervals.xlsx",
                   "04_LTE_BSS_Issus.xlsx",
                   "05_LTE_BSS_Issus_Intervals.xlsx"]

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




```python
#re-set all the variable from the RAM
%reset -f
```
