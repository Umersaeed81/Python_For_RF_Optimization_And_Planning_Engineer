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
import os
import streamlit as st
from datetime import datetime, timedelta
import win32com.client
import pythoncom
import time

# --- Streamlit Page Config ---
st.set_page_config(page_title="Outlook Attachment Downloader", layout="wide")
st.markdown("<h1 style='color: maroon;'>📥 Outlook Attachment Downloader</h1>", unsafe_allow_html=True)
# ------------------------------------------------------------------
# Initialize COM
pythoncom.CoInitialize()

# Default values
default_download_path = r"D:\Downloaded_Attachments"
today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
tomorrow = today + timedelta(days=1)

# PST Selection using checkboxes
st.markdown("<h4 style='color: maroon;'>📌 Select PST Option</h4>", unsafe_allow_html=True)

use_current_pst = st.checkbox("💼 Use current Outlook PST", value=False)
browse_pst_file = st.checkbox("🧭 Browse PST file", value=False)

root_folder = None

if use_current_pst and browse_pst_file:
    st.warning("⚠️ Please select only one PST option.")
elif not use_current_pst and not browse_pst_file:
    st.info("ℹ️ Please select a PST option to proceed.")
else:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    if browse_pst_file:
        pst_path = st.text_input("Enter full path to PST file (e.g., E:\\Email\\LTE_KPI_REPORTING.pst)")
        if pst_path:
            if not os.path.exists(pst_path):
                st.error("❌ PST file not found.")
            else:
                try:
                    outlook.AddStore(pst_path)
                    for i in range(outlook.Folders.Count):
                        folder = outlook.Folders.Item(i + 1)
                        if folder.FolderPath.endswith(os.path.basename(pst_path).replace(".pst", "")):
                            root_folder = folder
                            break
                    if not root_folder:
                        st.error("❌ Could not locate the root folder for the provided PST path.")
                except Exception as e:
                    st.error(f"❌ Error loading PST: {e}")
    elif use_current_pst:
        root_folder = outlook.GetDefaultFolder(6)

# Function to get all subfolders recursively
def get_all_folders(folder, path=""):
    full_path = os.path.join(path, folder.Name) if path else folder.Name
    folders = [(full_path, folder)]
    for sub in folder.Folders:
        folders.extend(get_all_folders(sub, full_path))
    return folders

if root_folder:
    all_folders = get_all_folders(root_folder)

    st.markdown("<h4 style='color: maroon;'>🗃️ Filter Folders for Attachment Extraction</h4>", unsafe_allow_html=True)
 
    st.markdown("> ✅ **Tip:** Select one or more folders from the list below where you expect the emails with attachments to be stored. These folders could include your Inbox, or any custom folders within the selected PST. Only selected folders will be scanned for matching emails.")
    selected_paths = []
    for folder_path, folder_obj in all_folders:
        if st.checkbox(folder_path, value=False):
            selected_paths.append((folder_path, folder_obj))

    if selected_paths:
        
        st.markdown("<h4 style='color: maroon;'>📧 Filter Emails by Sender</h4>", unsafe_allow_html=True)
        check_sender = st.checkbox("(optional) Filter by specific sender email?")
        sender_email = ""
        if check_sender:
            sender_email = st.text_input("Enter sender email (e.g., PRS.TOOL@ptclgroup.com)")

       
        st.markdown("<h4 style='color: maroon;'>🕵️ Filter Emails by Subject</h4>", unsafe_allow_html=True)
        check_subject = st.checkbox("(optional) Search Emails by Subject Text?")
        subject_keyword = ""
        if check_subject:
            subject_keyword = st.text_input("Subject must contain keyword", value="")

        #st.markdown("#### 🗓️ Filter Emails by Date:")
        st.markdown("<h4 style='color: maroon;'>🗓️ Filter Emails by Date</h4>", unsafe_allow_html=True)
        check_date = st.checkbox("(optional) Filter by date range?")
        if check_date:
            date_start_input = st.date_input("📅 Start Date", value=today.date())
            date_end_input = st.date_input("📅 End Date", value=tomorrow.date())
            date_start = datetime.combine(date_start_input, datetime.min.time())
            date_end = datetime.combine(date_end_input, datetime.min.time())
        else:
            date_start = datetime.min
            date_end = datetime.max

        
        st.markdown("<h4 style='color: maroon;'>🧩 Filter Attachments by File Type</h4>", unsafe_allow_html=True)
        check_ext = st.checkbox("(optional) Filter by attachment extension?")
        ext_filter = ""
        if check_ext:
            ext_filter = st.text_input("Only download attachments with this extension (e.g., .xlsx, .pdf)")

        #st.markdown("#### 📂 Select Download Location:")
        st.markdown("<h4 style='color: maroon;'>📂 Select Download Location</h4>", unsafe_allow_html=True)
        download_path = st.text_input("Download path", value=default_download_path)
        os.makedirs(download_path, exist_ok=True)

       # st.markdown("#### 🛠️ Download Mode Configuration:")
        st.markdown("<h4 style='color: maroon;'>🛠️ Download Mode Configuration</h4>", unsafe_allow_html=True)
        download_mode = st.selectbox("Choose download mode:", [
            "Delete all files if exist.",
            "Do not overwrite, rename it."
        ])

        if st.button("📥 Start Download"):

            # 🗑️ Delete all files at the start if selected
            if download_mode == "Delete all files if exist.":
                st.warning("🗑️ Deleting all existing files in the download folder...")
                for file in os.listdir(download_path):
                    file_path = os.path.join(download_path, file)
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                time.sleep(2)

            # 💡 Renaming logic
            def get_unique_filename(file_path):
                base, ext = os.path.splitext(file_path)
                counter = 1
                new_file_path = file_path
                while os.path.exists(new_file_path):
                    new_file_path = f"{base}_{counter}{ext}"
                    counter += 1
                return new_file_path

            def download_attachments(folders, sender_filter, subject_filter, date_start, date_end, extension_filter, download_path):
                file_count = 0
                for folder_path, folder_obj in folders:
                    try:
                        messages = folder_obj.Items
                        messages.Sort("[ReceivedTime]", True)

                        start_str = date_start.strftime('%m/%d/%Y %H:%M %p')
                        end_str = date_end.strftime('%m/%d/%Y %H:%M %p')
                        restricted_items = messages.Restrict(f"[ReceivedTime] >= '{start_str}' AND [ReceivedTime] < '{end_str}'")

                        index = 1
                        while index <= restricted_items.Count:
                            message = restricted_items.Item(index)
                            index += 1

                            if message.Class == 43:  # MailItem
                                sender = message.SenderEmailAddress
                                subject = message.Subject

                                if (not sender_filter or sender == sender_filter) and \
                                   (not subject_filter or subject_filter in subject):
                                    for i in range(1, message.Attachments.Count + 1):
                                        attachment = message.Attachments.Item(i)
                                        filename = attachment.FileName

                                        if extension_filter:
                                            if not filename.lower().endswith(extension_filter.lower()):
                                                continue

                                        save_path = os.path.join(download_path, filename)

                                        # ✅ Always rename if file already exists
                                        if os.path.exists(save_path):
                                            save_path = get_unique_filename(save_path)

                                        attachment.SaveAsFile(save_path)
                                        file_count += 1

                    except Exception as e:
                        st.warning(f"⚠️ Error in folder '{folder_path}': {e}")
                return file_count

            count = download_attachments(
                selected_paths,
                sender_email.strip() if check_sender else "",
                subject_keyword.strip() if check_subject else "",
                date_start,
                date_end,
                ext_filter.strip() if check_ext else "",
                download_path
            )

            st.success(f"✅ Download complete. Total attachments downloaded: {count}")


