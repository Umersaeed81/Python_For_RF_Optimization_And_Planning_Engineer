#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>


```python
import os

def run_script(script_path):
    return os.system(f'jupyter nbconvert --execute --to notebook --inplace {script_path}')

# List of script paths
script_paths = [
      "C:/Users/UWX161178/Daily_Audits_26052024/UMTS_LTE_Utilization/00_Folder_Management.ipynb",
       "C:/Users/UWX161178/Daily_Audits_26052024/UMTS_LTE_Utilization/01_UMTS_RF_Export_Processing.ipynb",
       "C:/Users/UWX161178/Daily_Audits_26052024/UMTS_LTE_Utilization/02_UMTS_Cell_List_From_PRS_KPIs.ipynb",  
       "C:/Users/UWX161178/Daily_Audits_26052024/UMTS_LTE_Utilization/03_UMTS_Cell_Mapping.ipynb",
      "C:/Users/UWX161178/Daily_Audits_26052024/UMTS_LTE_Utilization/04_UMTS_KPIs_Calculation.ipynb",
     "C:/Users/UWX161178/Daily_Audits_26052024/UMTS_LTE_Utilization/05_LTE_RF_Export_Processing.ipynb",
      "C:/Users/UWX161178/Daily_Audits_26052024/UMTS_LTE_Utilization/06_LTE_Cell_List_From_PRS_KPIs.ipynb",
      "C:/Users/UWX161178/Daily_Audits_26052024/UMTS_LTE_Utilization/07_LTE_Cell_Mapping.ipynb",
      "C:/Users/UWX161178/Daily_Audits_26052024/UMTS_LTE_Utilization/08_LTE_KPIs_Calculation.ipynb",
      "C:/Users/UWX161178/Daily_Audits_26052024/UMTS_LTE_Utilization/09_Project_Key.ipynb",
      "C:/Users/UWX161178/Daily_Audits_26052024/UMTS_LTE_Utilization/10_Project_Cell_Mapping.ipynb",
      "C:/Users/UWX161178/Daily_Audits_26052024/UMTS_LTE_Utilization/11_Merge_Mapping_with_KPIs.ipynb",
      "C:/Users/UWX161178/Daily_Audits_26052024/UMTS_LTE_Utilization/12_UMTS_LTE_Capacity_Analysis.ipynb"
]

error_scripts = []

for path in script_paths:
    return_code = run_script(path)
    if return_code != 0:
        error_scripts.append(path)
        print(f"Error in script: {path}")
        break  # Stop the loop if an error is encountered
    else:
        print(f"Script executed successfully: {path}")
        print("Done")

if error_scripts:
    print(f"The following scripts had errors:\n {error_scripts}")
```
