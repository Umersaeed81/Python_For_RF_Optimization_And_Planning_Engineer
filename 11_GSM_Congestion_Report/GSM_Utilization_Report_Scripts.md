#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>


```python
from datetime import datetime
start_time = datetime.now()
print(start_time)
```


```python
import os

def run_script(script_path):
    return os.system(f'jupyter nbconvert --execute --to notebook --inplace {script_path}')

script_paths = [
          "E:/Daily_Adit_Code/GSM_Congestion_Report/00_RF_Export(Cell_Level).ipynb",
          "E:/Daily_Adit_Code/GSM_Congestion_Report/01_RF_Export(GTRX).ipynb",
          "E:/Daily_Adit_Code/GSM_Congestion_Report/02_RF_Export(Frequency).ipynb",
           "E:/Daily_Adit_Code/GSM_Congestion_Report/03_PRS_Cluster_Defination.ipynb",
           "E:/Daily_Adit_Code/GSM_Congestion_Report/04_RF_Export(Final).ipynb",
           "E:/Daily_Adit_Code/GSM_Congestion_Report/05_BTS_Type_Identification.ipynb",
           "E:/Daily_Adit_Code/GSM_Congestion_Report/06_TXN_Type_Identificaton.ipynb",
           "E:/Daily_Adit_Code/GSM_Congestion_Report/07_TXN_BTS_Type_Identificaton.ipynb",
           "E:/Daily_Adit_Code/GSM_Congestion_Report/08_RF_Export_BTS_TXN_Type.ipynb",
            "E:/Daily_Adit_Code/GSM_Congestion_Report/09_GSM_BH_KPIs_Processing.ipynb",         
            "E:/Daily_Adit_Code/GSM_Congestion_Report/10_Excel_File_Formatting.ipynb"
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


```python
end_time = datetime.now()
print(end_time)
```
