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
              # Location  
              "C:/Users/UWX161178/Daily_Audits_26052024/WCMS/00_WCMS_Location.ipynb",
              # GSM KPIs  
              "C:/Users/UWX161178/Daily_Audits_26052024/WCMS/01_GSM_KPIs_WCMS.ipynb",    
              # LTE KPIs  
              "C:/Users/UWX161178/Daily_Audits_26052024/WCMS/02_LTE_KPIs_WCMS.ipynb", 
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
