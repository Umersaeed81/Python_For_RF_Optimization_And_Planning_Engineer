```python
import os
import win32com.client
```


```python
# Define PST file path and folder name
pst_path = r"E:\PRS_Email\LTE_KPI_REPORTING.pst"
folder_name = "Repots"
save_folder = r"E:\PRS_Email\Attachments"  

# Ensure the save folder exists
os.makedirs(save_folder, exist_ok=True)
```


```python
# Initialize Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
```


```python
# Add PST file if not already added
namespace.AddStore(pst_path)
```


```python
# Get the PST root folder
root_folder = namespace.Folders["LTE_KPI_REPORTING"]
```


```python
# Access the specific folder inside the PST
reports_folder = root_folder.Folders(folder_name)
```


```python
# Loop through all emails in the folder
for mail in reports_folder.Items:
    if mail.Attachments.Count > 0:
        for attachment in mail.Attachments:
            attachment_path = os.path.join(save_folder, attachment.FileName)
            attachment.SaveAsFile(attachment_path)
            print(f"Saved: {attachment_path}")

print("Attachment download complete.")
```
