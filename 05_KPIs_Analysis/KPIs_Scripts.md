#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

# SLA KPIs


```python
import os

def run_script(script_path):
    return os.system(f'jupyter nbconvert --execute --to notebook --inplace {script_path}')

# List of script paths
script_paths = [
                 "C:/Users/UWX161178/Daily_Audits_26052024/SLA_Conformance/00_Inter_BSC_HSR.ipynb",
                 "C:/Users/UWX161178/Daily_Audits_26052024/SLA_Conformance/01_2G_Daily_Conformance.ipynb",
                 "C:/Users/UWX161178/Daily_Audits_26052024/SLA_Conformance/04_2G_Week_Conformance_BH.ipynb",
                 "C:/Users/UWX161178/Daily_Audits_26052024/SLA_Conformance/03_2G_Month_Conformance_BH.ipynb",
                 "C:/Users/UWX161178/Daily_Audits_26052024/SLA_Conformance/02_2G_Quarterly_Conformance.ipynb",
                 "C:/Users/UWX161178/Daily_Audits_26052024/SLA_Conformance/05_Cluster_DA_Month_Week_Level_KPIs.ipynb"
]

# Execute scripts and track errors
for path in script_paths:
    return_code = run_script(path)
    if return_code == 0:
        print(f"Script {path} ran successfully.")
    else:
        print(f"Error: Script {path} encountered an error. Please recheck.")
```
