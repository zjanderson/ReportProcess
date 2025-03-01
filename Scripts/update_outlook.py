import win32com.client
import time

def access_inbox():
    try:
        # Create single Outlook instance outside the loop
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)

        explorer = outlook.ActiveExplorer()
        if explorer is None:
            explorer = outlook.Explorers.Add(inbox, 0)  # 0 is olFolderDisplayNormal
        explorer.Display()
        explorer.Activate()

        return inbox
    
    except Exception as e:
        print(f"Critical error in Outlook connection: {e}")

        return e
    
def force_update_inbox(inbox):
    try:
        # Force synchronization of the inbox
        inbox.Items.GetFirst()  # Get first item to trigger sync
        inbox.Items.GetLast()
        
        # Alternative sync method if needed
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        namespace.SendAndReceive(False)  
        time.sleep(100)
                
        print("Inbox updated successfully")
        return True
        
    except Exception as e:
        print(f"Error updating inbox: {e}")
        return False

    
if __name__ == "__main__":
    inbox = access_inbox()
    force_update_inbox(inbox)