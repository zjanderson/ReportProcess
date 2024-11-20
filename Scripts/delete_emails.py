import win32com.client

def delete_emails_from_folder(folder_path):
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
            
        # Delete from all subfolders
        for subfolder in current_folder.Folders:
            for item in list(subfolder.Items):
                item.Delete()
    else:
        # Original logic for specific folders
        folders = folder_path.split('/')
        current_folder = inbox
        for folder in folders:
            current_folder = current_folder.Folders.Item(folder)
        
        for item in list(current_folder.Items):
            item.Delete()


DELETE_FOLDERS = [
    "Deleted Items", 
    "Inbox/Weather Updates", 
    "Inbox/Fresh Beef", 
    "Inbox/Coverage", 
    "Inbox/BluePrism",
    "Inbox/National Accts/Chik Fil A/*"
    ]

def execute_deletes():
    for folder in DELETE_FOLDERS:
        print(folder)
        delete_emails_from_folder(folder)

if __name__ == "__main__":
    execute_deletes()