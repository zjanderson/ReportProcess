import win32com.client
import time
import sys
import os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from Supporting_Documents.credentials import MYEMAIL


DELETE_FOLDERS = [
    "Inbox/Weather Updates",
    "Inbox/Coverage",
    "Inbox/BluePrism",
    "Inbox/CFA 2.0",
    "Inbox/CFA Canada",
    "Inbox/CFA Hawaii",
    "Inbox/CFA Contingency",
    "Inbox/CFA Hormel",
    "Inbox/CFA Hubs",
    "Inbox/CFA PR",
    "Inbox/CFA FA",
    "Inbox/CFA MB",
    "Inbox/CFA McLane",
    "Inbox/CFA Perishables",
    "Inbox/CFA QCD",
    "Deleted Items",
]


def delete_app_emails_from_folder(folder_path):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.Folders.Item(MYEMAIL)

    # Handle wildcard pattern
    if "*" in folder_path:
        parent_path = folder_path.replace("/*", "")
        folders = parent_path.split("/")
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
        folders = folder_path.split("/")
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
