# [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)  
**Senior RF Planning & Optimization Engineer**  


📍 **Location:** Dream Gardens, Defence Road, Lahore  
📞 **Mobile:** +92 301 8412180  
✉ **Email:** [umersaeed81@hotmail.com](mailto:umersaeed81@hotmail.com)  

## **Education**  
🎓 **BSc Telecommunications Engineering** – [School of Engineering](https://sen.umt.edu.pk/)  
🎓 **MS Data Science** – [School of Business and Economics](https://hsm.umt.edu.pk/)  
**[University of Management & Technology](https://www.umt.edu.pk/)**  

-----------------------------------------------------------------------------------

# Filtering ZIP Files That Do Not Contain Specific Keywords

## Introduction

Handling compressed files efficiently is an essential skill for managing data in bulk. Often, we encounter scenarios where we need to process ZIP files while ensuring they meet specific conditions. One such case is filtering ZIP files based on their contents. This article discusses how to identify ZIP files that do not contain a particular keyword in any of their internal filenames.

## Understanding the Need for Filtering ZIP Files

ZIP files often contain multiple files, and in many cases, we need to analyze or extract only those that meet certain criteria. For example, an organization may receive daily reports in ZIP format, but only a subset of those reports is relevant to a specific analysis. Instead of extracting every file manually, we can automate the filtering process.

## Key Steps in the Filtering Process

To achieve this, the process involves:

1. **Locating the ZIP Files**: The first step is to specify the directory containing the ZIP files.

2. **Reading ZIP File Contents**: Each ZIP file is opened and scanned to retrieve the list of filenames it contains.

3. **Applying the Filtering Condition**: The filenames within each ZIP file are checked for the presence of a specific keyword.

4. **Listing the Filtered ZIP Files**: Finally, the ZIP files that do not contain the keyword are displayed.

## Real-World Applications

- **Automated Report Processing**: Filtering ZIP files that contain only the required reports can save time in data analysis.

- **Data Quality Checks**: Ensuring that required files exist within ZIP archives before further processing.

- **File Management & Cleanup**: Identifying unnecessary ZIP files that do not contain relevant data.

### 🏗 Import Required Modules


```python
import zipfile                        # 📦 Module for working with ZIP files  
from pathlib import Path              # 🛤 Module for handling file paths  
```

### 📂 Define the ZIP Folder


```python
# 📌 Set the path where ZIP files are stored  
zip_folder = Path(r"E:\PRS_Email\Attachments")  
```

### 🔍 Find All ZIP Files in the Folder


```python
# 📋 Get a list of all ZIP files in the specified folder  
zip_files = list(zip_folder.glob("*.zip"))  
```

### 🚀 Filter ZIP Files That Do Not Contain 'LTE_DA'


```python
# 📂 List to store ZIP files that do not contain 'LTE_DA' in any filename  
filtered_zips = []

for zip_file in zip_files:
    with zipfile.ZipFile(zip_file, 'r') as zf:
        # 📜 Get the list of file names inside the ZIP  
        file_names = zf.namelist()
        
        # 🔎 Check if any file name contains 'LTE_DA'  
        if not any("LTE_DA" in name for name in file_names):
            filtered_zips.append(zip_file.name)  # ✅ Add ZIP file to the filtered list  
```

### 📢 Display Filtered ZIP Files


```python
# 🖨 Print ZIP files that do not contain any file with 'LTE_DA'  
print("ZIP files that do not contain 'LTE_DA':")
for zip_name in filtered_zips:
    print(zip_name)  # 📜 Display the ZIP filenames  
```

    ZIP files that do not contain 'LTE_DA':
    

## Conclusion

By leveraging automation, filtering ZIP files based on their contents becomes a streamlined process, eliminating the need for manual inspections. This method can be particularly useful in handling large volumes of data efficiently.

For a complete implementation, you can check out the code on GitHub: [[GitHub Repository Link](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/pst_attachmets_working/01_Filtering_ZIP_Files_That_Do_Not_Contain_Specific_Keywords.md)]
