#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# GSM CDIG


```python
import os

def run_script(script_path):
    return os.system(f'jupyter nbconvert --execute --to notebook --inplace {script_path}')

# List of script paths
script_paths = [
#     SITE LEVEL AVAILABILITY IN PIVOT TABLE (LAST 7 DAYS)
    "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/13_GSM_CGID_Logs_Analysis_A.ipynb",
#-----------------------------------------------------------------------------------------------------#
#   Have to set the number of day day before run the following code, normall on weekend
    "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/14_GSM_CGID_Logs_Analysis_B.ipynb",
#-----------------------------------------------------------------------------------------------------#
    
    
#     #lAst 7 DAYS OUTAGE INTERVAL IN PIVOT TABLE FORMAT
       "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/15_GSM_CGID_Logs_Analysis_C.ipynb",
   
    #CDIG FILE PROCESSING
    "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/16_GSM_CGID_Logs_Analysis_D.ipynb",
    
    #COMBINE ALL EXCEL FILES
    "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/17_GSM_CGID_Logs_Analysis_E.ipynb",
    
    #FORMAT EXCEL FILE
    "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/18_GSM_CGID_Logs_Analysis_F.ipynb"
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
