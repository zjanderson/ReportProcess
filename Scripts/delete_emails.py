import win32com.client
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.common.exceptions import TimeoutException
import time


DELETE_FOLDERS = [
    "Inbox/Weather Updates", 
    "Inbox/Fresh Beef", 
    "Inbox/Coverage", 
    "Inbox/BluePrism",
    "Inbox/National Accts/Chik Fil A/*",
    "Deleted Items", 
    ]

FOLDER_TO_CLASS = {
    "Inbox/Weather Updates" : "gtcPn _8g73 LPIso", 
    "Inbox/Fresh Beef" : "", 
    "Inbox/Coverage" : "", 
    "Inbox/BluePrism" : "",
    "Inbox/National Accts/Chik Fil A/*" : "",
    "Inbox/National Accts/Chik Fil A/*" : "",
    "Inbox/National Accts/Chik Fil A/*" : "",
    "Inbox/National Accts/Chik Fil A/*" : "",
    "Inbox/National Accts/Chik Fil A/*" : "",
    "Inbox/National Accts/Chik Fil A/*" : "",
    "Inbox/National Accts/Chik Fil A/*" : "",
    "Inbox/National Accts/Chik Fil A/*" : "",
    "Inbox/National Accts/Chik Fil A/*" : "",
    "Inbox/National Accts/Chik Fil A/*" : "",
    "Deleted Items" : "", 
}

def delete_app_emails_from_folder(folder_path):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.Folders.Item("zanderson@armada.net")
    
    # Handle wildcard pattern
    if '*' in folder_path:
        parent_path = folder_path.replace('/*', '')
        folders = parent_path.split('/')
        current_folder = inbox
        
        # Navigate to parent folder
        for folder in folders:
            current_folder = current_folder.Folders.Item(folder)
            
        # Empty all subfolders
        for subfolder in current_folder.Folders:
            while subfolder.Items.Count > 0:
                # Delete each item individually
                subfolder.Items.Item(1).Delete()
    else:
        # Original logic for specific folders
        folders = folder_path.split('/')
        current_folder = inbox
        for folder in folders:
            current_folder = current_folder.Folders.Item(folder)
        
        # Delete items one by one instead of using Clear()
        while current_folder.Items.Count > 0:
            current_folder.Items.Item(1).Delete()

def execute_app_deletes():
    for folder in DELETE_FOLDERS:
        print(folder)
        delete_app_emails_from_folder(folder)



if __name__ == "__main__":
    execute_app_deletes()
