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
               
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/00_2G_EI_DA_Average_Calculation.ipynb",
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/01_3G_EI_DA_Average_Calculation.ipynb",
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/02_4G_EI_DA_Average_Calculation.ipynb",
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/03_GSM_Interval_In_Pivot_Table.ipynb",
    
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/04_UMTS_Interval_In_Pivot_Table.ipynb",
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/05_LTE_Interval_In_Pivot_Table.ipynb",
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/06_GSM_Max_IOI_In_Pivot_Table.ipynb",
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/07_UMTS_Max_IOI_In_Pivot_Table.ipynb",
    
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/08_LTE_Max_IOI_In_Pivot_Table.ipynb",
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/09_GSM_EI_Process_File_Final.ipynb",
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/10_UMTS_EI_Process_File_Final.ipynb",
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/11_LTE_EI_Process_File_Final.ipynb",
    
            
             "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/12_Del_Unwanted_Files.ipynb",
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/13_Combined_Excel_File.ipynb",
              "C:/Users/UWX161178/Daily_Audits_26052024/EI_Report/14_External_Interference_Formatting.ipynb",
             
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
