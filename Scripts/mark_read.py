import win32com.client
from datetime import time, datetime
import sys


ALL_FOLDERS = [
    # "IB Hub Dallas", 
    # "IB Hub East Point", 
    # "IB Hub Greencastle", 
    # "IB Hub Romeoville", 
    # "MCDToys", 
    # "MCD East", 
    # "MCD South", 
    # "MCD Central", 
    # "MCD West", 
    # "MCD Supply", 
    # "Zaxby's", 
    # "Bojangles", 
    # "Stakeholders", 
    # "Supply Caddy", 
    # "BBI", 
    # "CFA Canada",
    # "CFA Hawaii",
    # "CFA Contingency",
    # "CFA Hormel",
    # "CFA Hubs",
    # "CFA PR",
    # "CFA 2.0",
    # "CFA FA",
    # "CFA MB",
    # "CFA McLane",
    # "CFA Perishables",
    # "CFA QCD",
    # "Darden", 
    "Darden/DDL", 
    # "Darden/DDL Maines", 
    # "Darden/DDL McLane", 
    # "Dominoes", 
    # "Panda Express",
    # "Panda Produce", 
    # "Panera", 
    # "Panera Chips", 
    # "Panera PandaEx GFS", 
    # "Panera PandaEx SYGMA", 
    # "QA", 
    # "Fresh Beef",
    # "Weather Updates"
    ]

def access_inbox():
    try:
        # Create single Outlook instance outside the loop
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)

        return inbox
    
    except Exception as e:
        print(f"Critical error in Outlook connection: {e}")

        return e
    
def mark_emails_in_folder_read(folder, military_time):
    try:
        today = datetime.now().date()
        cutoff_time = datetime.combine(today, time(hour=military_time))
        print(f"Cutoff time: {cutoff_time}")
        
        # Get all unread items in the folder
        items = folder.Items.Restrict("[Unread] = True")
        print(f"Found {len(items)} unread emails in {folder.Name}.. marking ones before {military_time}:00 as read")
        
        count = 0
        for item in items:
            received_time = item.ReceivedTime
            received_datetime = datetime.combine(received_time.date(), received_time.time())
            if received_datetime < cutoff_time:
                item.UnRead = False
                item.Save()
                count += 1

        print(f"Marked {count} emails as read in folder: {folder.Name}")
        
    except Exception as e:
        print(f"Error processing folder {folder.Name}: {e}")


def process_folders(time):
    inbox = access_inbox()
    for folder_name in ALL_FOLDERS:
        folder = inbox.Folders.Item(folder_name)
        mark_emails_in_folder_read(folder, time)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python mark_read.py <military_time>")
        print("Example: python mark_read.py 14")
        sys.exit(1)
    
    try:
        military_time = int(sys.argv[1])
        if 0 <= military_time <= 23:
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)
            process_folders(military_time)

        else:
            print("Error: Please enter a valid hour between 0 and 23")
            sys.exit(1)
    except ValueError:
        print("Error: Please enter a valid integer")
        sys.exit(1)

