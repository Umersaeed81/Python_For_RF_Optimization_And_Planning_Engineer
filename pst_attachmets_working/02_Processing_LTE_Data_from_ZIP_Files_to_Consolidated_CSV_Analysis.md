# [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)  
**Senior RF Planning & Optimization Engineer**  


📍 **Location:** Dream Gardens, Defence Road, Lahore  
📞 **Mobile:** +92 301 8412180  
✉ **Email:** [umersaeed81@hotmail.com](mailto:umersaeed81@hotmail.com)  

## **Education**  
🎓 **BSc Telecommunications Engineering** – School of Engineering  
🎓 **MS Data Science** – School of Business and Economics  
**University of Management & Technology**  

-----------------------------------------------------------------------------------

# 📂 Processing LTE Data from ZIP Files to Consolidated CSV Analysis 📊

This script processes **multiple ZIP files**, each containing **multiple CSV files**. It scans all ZIP files in the specified folder and extracts only those CSV files that have "**LTE_DA**" in their filename. The goal is to **read specific CSV files**, clean and format the extracted data, and merge it with a site list to ensure only relevant records are included. The data is then summarized, analyzed, and used to generate a **final report** for performance evaluation. The processed results are saved as an **Excel file** for further analysis.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/pst_attachmets_working/PIC_06n.png?raw=true)

### 🗂️ Import Required Libraries


```python
import os                         # 📁 Handle file system operations  
import zipfile                    # 📦 Work with ZIP files  
import numpy as np                # 🔢 Numerical computations  
import pandas as pd               # 📊 Data manipulation  
from pathlib import Path          # 🛤️ Handle file paths  
```

- `os` → Helps with file and directory handling 📂
- `zipfile` → Extracts data from ZIP archives 📦
- `numpy` (np) → Used for numerical operations 🔢
- `pandas` (pd) → Processes CSV & Excel data 📊
- `pathlib.Path` → Helps find ZIP files in a folder 📁

### 📂 Define ZIP Folder Path


```python
zip_folder = "E:/PRS_Email/Attachments"  # 📌 Set the directory containing ZIP files  
```

- This is the folder where all ZIP files are stored 📂

### 🔍 Retrieve ZIP Files


```python
# 📄 Get a list of all ZIP files in the specified folder  
zip_files = list(Path(zip_folder).glob("*.zip"))  
```

- This finds all `.zip` files inside the folder 📦

### 📥 Extract and Read CSV Files from ZIP Archives


```python
dfs = []  # 📝 Initialize an empty list to store DataFrames  

# 🎯 Define the required columns  
required_columns = ['Date','Cell Name','eNodeB Function Name','Frequency band',
                    'Downlink EARFCN','Data Volume (GB)','DL User Thrp (Mbps)_DEN',
                    'DL User Thrp (Mbps)_NUM']

# 🔄 Loop through each ZIP file  
for zip_file in zip_files:
    with zipfile.ZipFile(zip_file, 'r') as z:
        # 📃 List CSV files containing "LTE_DA"  
        csv_files = [f for f in z.namelist() if "LTE_DA" in f and f.endswith(".csv")]
        
        for csv_file in csv_files:
            with z.open(csv_file) as f:
                try:
                    # 📊 Read required columns, parse 'Date' as datetime  
                    df = pd.read_csv(f, usecols=required_columns, parse_dates=["Date"],skiprows=range(6),\
                                skipfooter=1,engine='python',na_values=['NIL','/0'])
                    df["source_file"] = csv_file  # 🏷️ Add file name tracking  
                    df["zip_file"] = zip_file.name  # 🏷️ Add ZIP file tracking  
                    dfs.append(df)  # 📌 Store DataFrame  
                except ValueError:  
                    # 🚨 Handle missing columns  
                    print(f"Skipping {csv_file} in {zip_file.name} - Missing required columns.")

# 📌 Combine all extracted data into one DataFrame  
if dfs:
    final_df = pd.concat(dfs, ignore_index=True)  # 🔄 Merge DataFrames  
else:
    print("No matching CSV files found or required columns missing.")  # ❌ Show error message  
```

- 📄 **Reads CSV files from each ZIP** to extract relevant data.
- 🧐 **Filters files containing "LTE_DA"** to focus on relevant datasets.
- 📊 **Extracts only the required columns** for efficient processing.
- ❌ Handles missing values (`NIL`, `/0`) by converting them to `NaN`.
- 📋 **Merges all extracted CSV data into a single table** (f`inal_df`).
- ✅ Skips footer rows using `skipfooter=1` to exclude unwanted summary rows.
- 🏷️ **Creates a 'source_file' column** to keep track of the original CSV file name.
- 🏷️ **Creates a 'zip_file' column** to record the source ZIP file for each dataset.
- 🚨 **Notifies the user** when a file is skipped due to missing required columns, explicitly mentioning the affected file name.
- 📅 **Parses 'Date' column** as a datetime format to ensure consistency in date-related operations.
- 🛠 Uses `engine=python` to handle complex file structures, such as skipping rows and footers.
- 🔄 **Concatenates all DataFrames** into a single DataFrame (`final_df`) for easy analysis.
- ❌ **Displays an error message** when no valid CSV files are found or all files lack required columns.

### ✏️ Clean Column Values


```python
# 🧹 Remove extra spaces in 'eNodeB Function Name'  
final_df['eNodeB Function Name'] = final_df['eNodeB Function Name'].str.replace(r'\s+', ' ', regex=True)

# 🧹 Remove extra spaces in 'Cell Name'  
final_df['Cell Name'] = final_df['Cell Name'].str.replace(r'\s+', ' ', regex=True)
```

- Replaces multiple spaces with a single space in text fields 🧹

### 🏷️ Generate Site and Cell IDs


```python
# 🏗️ Extract LTE Site ID from 'eNodeB Function Name'  
final_df['LTE_Site_ID'] = final_df['eNodeB Function Name'].str.split('_').str[0]

# 🏗️ Extract LTE Cell ID from 'Cell Name'  
final_df['LTE_Cell_ID'] = final_df['Cell Name'].str.split('_').str[0]
```

- Generates unique Site ID and Cell ID from names 📌

### 📂 Set Working Directory


```python
path = 'E:/PRS_Email/req_sites'  # 📌 Define path  
os.chdir(path)  # 🔄 Change current working directory  
```

- Moves to the folder where we will process the data 📂

### 📑 Load Required Site List


```python
# 📥 Read the site list Excel file  
req_site_list = pd.read_excel('Site_List.xlsx')
```

- Loads **site reference data** from an Excel file 📖

### 🔄 Merge Site List with Extracted Data


```python
# 🔗 Merge required site list with extracted dataset  
data_set = pd.merge(req_site_list,final_df,how='left',on=['LTE_Site_ID'])
```

- Keeps only the required sites 📍
- Joins data using LTE_Site_ID as the key 🔑

### 📊 Aggregate KPI Metrics per Day


```python
# 📅 Group by 'Date', 'LTE_Site_ID', 'Province', 'Site integration Date' and calculate sum of KPIs  
col_sum_kpis_day= data_set.groupby(['Date','LTE_Site_ID','Province','Site integration Date'])\
        [['Data Volume (GB)',\
          'DL User Thrp (Mbps)_DEN','DL User Thrp (Mbps)_NUM']]\
        .sum().reset_index()
```

- Group data by day & site 📆
- Sums important KPIs 📊

### 📏Calculate Downlink User Throughput


```python
# 📈 Compute 'DL User Thrp (Mbps)'  
col_sum_kpis_day['DL User Thrp (Mbps)'] =  (col_sum_kpis_day['DL User Thrp (Mbps)_NUM'] / col_sum_kpis_day['DL User Thrp (Mbps)_DEN']) / 1024
```

- Computes throughput in Mbps 🔢

### 📆 Generate Week Identifier (WK)


```python
# 🗓️ Create 'WK' column for Year_Week format  
col_sum_kpis_day['WK'] = col_sum_kpis_day['Date'].dt.isocalendar().year.astype(str) + "_" + \
           col_sum_kpis_day['Date'].dt.isocalendar().week.astype(str).apply(lambda x: x.zfill(2))
```

- Creates Week labels for analysis 📅

### 📅 Generate Month Identifier


```python
# 📆 Create 'Month' column for Year_Month format  
col_sum_kpis_day['Month'] = col_sum_kpis_day['Date'].dt.year.astype(str) + "_" + \
               col_sum_kpis_day['Date'].dt.month.astype(str).apply(lambda x: x.zfill(2))
```

- Creates Month labels for analysis 📅

### 🚦Identify Sites Below KPI Threshold


```python
# 🚨 Check if 'DL User Thrp (Mbps)' is below 4 and mark as 'Yes' or 'No'  
col_sum_kpis_day['DL User Thrp (Mbps)<4(Yes/No)'] = np.where(
                    (col_sum_kpis_day['DL User Thrp (Mbps)'] < 4),
                    'Yes', 'No')
```

- Flags sites with low throughput (<4 Mbps) 🚨

### 🎯 Select Final Columns for Output


```python
# 📌 Keep only necessary columns for final dataset  
col_sum_kpis_day_final = col_sum_kpis_day[['Date','WK','Month','LTE_Site_ID', 'Province',
                        'Site integration Date','Data Volume (GB)', 'DL User Thrp (Mbps)_DEN',
       'DL User Thrp (Mbps)_NUM', 'DL User Thrp (Mbps)','DL User Thrp (Mbps)<4(Yes/No)']]
```

- Selects only **important columns** for reporting 📊

### 💾 Export Final Data to Excel


```python
col_sum_kpis_day_final.to_excel('DL_User_Thrp.xlsx',index=False)
```

- Exports final report to Excel 📤
