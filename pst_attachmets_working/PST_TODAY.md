# [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)  
**Senior RF Planning & Optimization Engineer**  


📍 **Location:** Dream Gardens, Defence Road, Lahore  
📞 **Mobile:** +92 301 8412180  
✉ **Email:** [umersaeed81@hotmail.com](mailto:umersaeed81@hotmail.com)  

## **Education**  
🎓 **BSc Telecommunications Engineering** – School of Engineering  
🎓 **MS Data Science** – School of Business and Economics  
**University of Management & Technology** 

------------------------------------------

# 📥📆 Automatically Download Today’s Outlook Email Attachments 📨⚡

## 📦 Import Required Libraries

Essential modules for date handling, file operations, and Outlook automation.


```python
import os
import win32com.client
from datetime import datetime, timedelta
```

## ⏰ Define Start and End Time for Today

Set today's date range from midnight to 11:59 PM for filtering emails.


```python
# Define today's start and end time
today = datetime.now()
start_date = datetime(today.year, today.month, today.day, 0, 0, 0)
end_date = datetime(today.year, today.month, today.day, 23, 59, 59)
```

## 📁 Create Folder to Save Attachments

Ensure the download folder exists for saving email attachments.


```python
# Set your local download folder
download_folder = r"D:\Downloaded_Attachments"  # Change this path as needed
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
# Sort and filter messages by today's date
messages = prs_folder.Items
messages.Sort("[ReceivedTime]", True)
```

## 🔍 Filter Emails by Date Range

Use DASL query syntax to restrict emails within the date range.


```python
# Restrict by today's date range using DASL syntax
restriction = f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y %H:%M %p')}' AND [ReceivedTime] <= '{end_date.strftime('%m/%d/%Y %H:%M %p')}'"
filtered_messages = messages.Restrict(restriction)
```

## 📎 Download Attachments from Filtered Emails

Loop through filtered emails and save all attachments to the folder.


```python
# Download attachments from filtered emails
for message in filtered_messages:
    if message.Class == 43:  # Ensure it's a MailItem
        for i in range(1, message.Attachments.Count + 1):
            attachment = message.Attachments.Item(i)
            file_path = os.path.join(download_folder, attachment.FileName)
            attachment.SaveAsFile(file_path)
            print(f"Downloaded: {attachment.FileName}")

print("Today's email attachments downloaded.")
```
