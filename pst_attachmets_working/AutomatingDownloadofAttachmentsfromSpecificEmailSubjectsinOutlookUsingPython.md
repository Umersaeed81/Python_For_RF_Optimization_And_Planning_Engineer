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

# 📥 Automating Download of Attachments from Specific Email Subjects in Outlook Using Python

## Import Required Libraries 📦🐍


```python
import os
import win32com.client
```

## Define Paths and Ensure Save Folder Exists 📁✅


```python
# Set download path
download_folder = r'D:\Downloaded_Attachments'  # Change this to your desired location
os.makedirs(download_folder, exist_ok=True)
```

## Initialize Outlook 💻📬


```python
# Connect to Outlook
outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
```

## 📂🔍 Navigating to the 'PRS' Folder Inside Outlook Inbox


```python
# Access the Inbox > PRS folder
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
prs_folder = None
for folder in inbox.Folders:
    if folder.Name == 'PRS':
        prs_folder = folder
        break

if not prs_folder:
    print("PRS folder not found in Inbox.")
    exit()
```

## 📬🎯 Filtering Emails by Subject and Downloading Attachments 📎💾


```python
# Filter emails by subject
messages = prs_folder.Items
messages = messages.Restrict("[Subject] = 'AILY_BSSP_3G_59'")

# Loop through filtered emails and download attachments
for message in messages:
    if message.Class == 43:  # Ensure it's a MailItem
        attachments = message.Attachments
        for i in range(1, attachments.Count + 1):
            attachment = attachments.Item(i)
            attachment.SaveAsFile(os.path.join(download_folder, attachment.FileName))
            print(f"Downloaded: {attachment.FileName}")

print("All matching attachments downloaded.")
```

    Downloaded: DAILY_BSSP_3G_59-20250303074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250302074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250301074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250304074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250305074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250306074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250307074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250310074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250309074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250308074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250311074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250312074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250313074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250314074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250317074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250316074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250315074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250318074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250319074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250320074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250321074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250323074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250322074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250324074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250325074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250326074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250327074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250328074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250401074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250331074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250330074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250329074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250406074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250405074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250404074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250403074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250402074854.zip
    Downloaded: DAILY_BSSP_3G_59-20250407074854.zip
    All matching attachments downloaded.
    


```python

```
