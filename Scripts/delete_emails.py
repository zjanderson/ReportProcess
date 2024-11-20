import win32com.client

def delete_emails_from_folder(folder_name):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.Folders.Item("zanderson@armada.net")
    folder = inbox.Folders.Item(folder_name)
   
    for item in list(folder.Items):
        item.Delete()

DELETE_FOLDERS = [
    "Deleted Items", 
    "Weather Updates", 
    "QCD", 
    "Perishables", 
    "McLane CFA", 
    "MB CFA", 
    "FA CFA", 
    "CFA PR", 
    "CFA Hubs", 
    "CFA Hormel", 
    "CFA Contingency", 
    "CFA Hawaii", 
    "CFA Canada", 
    "2.0 CFA", 
    "Fresh Beef", 
    "Coverage", 
    "BluePrism"
    ]

def execute_deletes():
    for folder in DELETE_FOLDERS:
        delete_emails_from_folder(folder)

if __name__ == "__main__":
    execute_deletes()