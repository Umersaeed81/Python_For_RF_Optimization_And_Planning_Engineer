## 📥📅 Download Outlook Email Attachments Within a Specific Date Range 📨📂

## 📦 Import Required Libraries

Essential modules for date handling, file operations, and Outlook automation.


```python
import os
import win32com.client
from datetime import datetime
```

## 📅 Set Date Range for Email Filtering

Define the start and end date for filtering emails.


```python
# Define the date range
start_date = datetime(2025, 4, 1)
end_date = datetime(2025, 4, 3, 23, 59, 59)  # Include entire end date
```

## 📁 Create Folder to Save Attachments

Ensure the download folder exists for saving email attachments.


```python
# Define where to save attachments
download_folder = r"D:\Downloaded_Attachments"  # Change path as needed
os.makedirs(download_folder, exist_ok=True)
```

## 📧 Connect to Outlook Application

Establish connection to Microsoft Outlook using COM.


```python
# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
```

## 📂 Access 'PRS' Folder in Inbox

Navigate to the 'PRS' subfolder inside the Inbox.


```python
# Access the Inbox > PRS folder
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
prs_folder = None
for folder in inbox.Folders:
    if folder.Name == 'PRS':
        prs_folder = folder
        break

if not prs_folder:
    print("PRS folder not found.")
    exit()
```

## 📨 Sort Emails by Received Time

Organize messages in descending order by received time.


```python
# Filter emails by date range
messages = prs_folder.Items
messages.Sort("[ReceivedTime]", True)  # Sort newest to oldest
```

## 🔍 Filter Emails by Date Range

Use DASL query syntax to restrict emails within the date range.


```python
# Restrict by date using DASL syntax
restriction = f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y %H:%M %p')}' AND [ReceivedTime] <= '{end_date.strftime('%m/%d/%Y %H:%M %p')}'"
filtered_messages = messages.Restrict(restriction)
```

## 📎 Download Attachments from Filtered Emails

Loop through filtered emails and save all attachments to the folder.


```python
# Download attachments
for message in filtered_messages:
    if message.Class == 43:  # Ensure it's a MailItem
        for i in range(1, message.Attachments.Count + 1):
            attachment = message.Attachments.Item(i)
            save_path = os.path.join(download_folder, attachment.FileName)
            attachment.SaveAsFile(save_path)
            print(f"Downloaded: {attachment.FileName}")

print("Attachments from selected date range downloaded.")
```
