#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# BSS KPIs


```python
import os

def run_script(script_path):
    return os.system(f'jupyter nbconvert --execute --to notebook --inplace {script_path}')

# List of script paths
script_paths = [
    
#--------------------------------------------------------------------------------------------------# 
##              # Have to set the number of day day before run the following code, normall on weekend
                "C:/Users/UWX161178/Daily_Audits_26052024/BSS/00_GSM_EI_Availability_Hourly.ipynb",
##              # Have to set the number of day day before run the following code, normall on weekend  
                "C:/Users/UWX161178/Daily_Audits_26052024/BSS/01_UMTS_EI_Availability_Hourly.ipynb",
##              # Have to set the number of day day before run the following code, normall on weekend 
                "C:/Users/UWX161178/Daily_Audits_26052024/BSS/02_LTE_EI_Availability_Hourly.ipynb",
##--------------------------------------------------------------------------------------------------#            
#                # Re-shape (pivot_table:GSM) 
                "C:/Users/UWX161178/Daily_Audits_26052024/BSS/06_GSM_Availability_Rate_Interval_7days.ipynb",
#                # Re-shape (pivot_table:UMTS)  
                "C:/Users/UWX161178/Daily_Audits_26052024/BSS/07_UMTS_Down_Time_Interval_7days.ipynb",
#                # Re-shape (pivot_table:LTE) 
                "C:/Users/UWX161178/Daily_Audits_26052024/BSS/08_LTE_Down_Time_Interval_7days.ipynb",
              
#                # Last 7 days Data Required: GSM  
                "C:/Users/UWX161178/Daily_Audits_26052024/BSS/03_GSM_Cell_DA.ipynb",
#                # Last 7 days Data Required: UMTS  
                "C:/Users/UWX161178/Daily_Audits_26052024/BSS/04_UMTS_Cell_DA.ipynb",
#                # Last 7 days Data Required: LTE
                "C:/Users/UWX161178/Daily_Audits_26052024/BSS/05_LTE_Cell_DA.ipynb",
              
               # Prossed Files  
               "C:/Users/UWX161178/Daily_Audits_26052024/BSS/09_UMTS_BSS.ipynb",
               # Del unwanted Files  
               "C:/Users/UWX161178/Daily_Audits_26052024/BSS/10_Del_unwanted_Files.ipynb",
              # Combine All excel Files  
              "C:/Users/UWX161178/Daily_Audits_26052024/BSS/11_Combined_Excel_File.ipynb",
              # Format Excel File  
              "C:/Users/UWX161178/Daily_Audits_26052024/BSS/12_Excel_File_Formatting.ipynb"              
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
