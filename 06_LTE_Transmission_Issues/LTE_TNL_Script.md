#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

## LTE TNL


```python
import os

def run_script(script_path):
    return os.system(f'jupyter nbconvert --execute --to notebook --inplace {script_path}')

# List of script paths
script_paths = [
     # Site level last 7 days data.
     "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/00_TNL_DA.ipynb",
     # Cell Level last 7 days data
     "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/01_Flash_CSFB.ipynb",
     # calculate last 7 days capacity
     "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/02_Evaluate_the_Transmission_Link_Capacity.ipynb",
    
     # SON Audit
     "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/08_SON_Audit.ipynb",
    
#----------------------------------------------------------------------------------------------------------------#    
    # Have to set the number of day day before run the following code, normall on weekend
    "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/03_Pure_TNL_0.ipynb",
#----------------------------------------------------------------------------------------------------------------#     
    
    # re-share data (pivot_table) last 7 days data
     "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/04_Pure_TNL_1.ipynb",
     # combine Base_Station_Transport (RF Export)
     "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/05_Base_Station_Transport_Data.ipynb",
 
#----------------------------------------------------------------------------------------------------------------#      
    # Have to set the number of day day before run the following code, normall on weekend
   "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/06_Max_User_Calculation.ipynb",
#----------------------------------------------------------------------------------------------------------------#    
    
#    Combine Licence, Max User re-shape(pivot table)
   "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/07_Max_User_Row_Lic_File_Processsing.ipynb",
     
     # file processing
     "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/10_TNL_Processing_File.ipynb",
     # TNL File Formatting
     "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/11_TNL_Excel_File_Formatting.ipynb",
     # Licence File Formatting
     "C:/Users/UWX161178/Daily_Audits_26052024/Transmission/12_LTE_NodB_Level_Licence_Formating.ipynb",

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
