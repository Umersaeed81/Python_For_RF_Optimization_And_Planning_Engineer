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

