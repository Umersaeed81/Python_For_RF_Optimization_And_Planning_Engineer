# 📨 Automating Outlook Attachment Downloads Based on Subject and Date using Python 🐍

## 🧾 Purpose of the Code

This Python script automates the process of downloading attachments from Microsoft Outlook emails based on two key conditions:
1. **The email was received today** (from midnight to just before midnight)
2. **The subject line contains a specific keyword** (e.g., daily reports or project files)

It navigates through the **entire Outlook Inbox**, including all subfolders, looking for matching emails. When it finds one, it saves all attachments from that message to a designated download folder on your system (e.g., D:\Downloaded_Attachments or any path you define).

Highlights of the script:

- 🔁 Recursively searches through all folders within the Inbox.
- 🗓️ Filters only emails received **today**.
- 🔤 Matches **keywords** in subject lines (case-sensitive).
- 📂 Downloads attachments to a specified local folder (you choose the location)
- ✅ Automatically creates the folder if it doesn't already exist
- ⚠️ Includes basic error handling to report issues per folder

This is ideal for automating tasks like downloading daily status reports, automated logs, or other recurring files that arrive in email.

## 📦 Import Required Libraries 📚


```python
import os                                  # 📁 Used for file and folder operations (like creating folders and joining paths)
import win32com.client                     # 💼 Used to interact with Outlook via COM interface (for accessing emails)
from datetime import datetime, timedelta   # 🕒 Used to handle and compare email dates
```

## 📂 Set Download Directory 📥


```python
download_folder = r'D:\Downloaded_Attachments'
os.makedirs(download_folder, exist_ok=True)  # ✅ Ensure the folder exists (create if not)
```

# 📬 Connect to Outlook Inbox 📧


```python
outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')  # 🔗 Establish Outlook connection
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox 📨
```

## 📅 Define Today's Date Range 🕓


```python
today_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)  # 🕛 Start of today
tomorrow_start = today_start + timedelta(days=1)  # 🕛 Start of tomorrow
```

## 🧮 Format Date & Set Target Keyword 🔍


```python
start_str = today_start.strftime('%m/%d/%Y %H:%M %p')  # 📆 Convert to string format for filtering
end_str = tomorrow_start.strftime('%m/%d/%Y %H:%M %p')
target_keyword = 'AILY_BSSP_3G_59'  # 🎯 Keyword to search in subject line
```

## 🔎 Recursive Folder Search & Attachment Download 💾


```python
def search_folder(folder):
    try:
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)  # 🔽 Sort messages by latest received time

        # 📌 Restrict to today's emails only
        date_filter = f"[ReceivedTime] >= '{start_str}' AND [ReceivedTime] < '{end_str}'"
        messages = messages.Restrict(date_filter)

        for message in messages:
            if message.Class == 43:  # ✅ Ensure it's a MailItem
                if target_keyword in message.Subject:  # 🧠 Check if subject contains the keyword
                    for i in range(1, message.Attachments.Count + 1):
                        attachment = message.Attachments.Item(i)
                        save_path = os.path.join(download_folder, attachment.FileName)
                        attachment.SaveAsFile(save_path)  # 💾 Save the attachment to disk
                        print(f"Downloaded from '{folder.Name}': {attachment.FileName}")
    except Exception as e:
        print(f"Error in folder '{folder.Name}': {str(e)}")  # ⚠️ Print any errors

    # 🔁 Recurse into subfolders
    for subfolder in folder.Folders:
        search_folder(subfolder)

# 🚀 Start Searching from Inbox
search_folder(inbox)
```

    Downloaded from 'Inbox': DAILY_BSSP_3G_59-20250418074854.zip
    Downloaded from 'PRS': DAILY_BSSP_3G_59-20250418074854.zip
    
