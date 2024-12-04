import win32com.client
import re
import nltk
from nltk.tokenize import word_tokenize

try:
    nltk.download('punkt')
except Exception as e:
    print(f"Error downloading NLTK data: {e}")


def extract_numbers(text):
    """
    Extract sequences of at least 5 numerical digits from the text.
    """
    numbers = re.findall(r'\b\d{5,}\b', text)
    return set(numbers)

def get_favorites_folders(outlook):
    """
    Get folders currently marked as 'Favorites' in Outlook's navigation pane.
    """
    try:
        # Access the Navigation Pane
        nav_pane = outlook.ActiveExplorer().NavigationPane
        # Access the Favorites navigation module
        nav_module = nav_pane.Modules.GetNavigationModule(0)  # 0 is olModuleMail
        # Get the actual Favorites folder group
        favorites_group = nav_module.NavigationGroups('Favorites')
        
        print("\nFavorites folders found:")
        favorites = []
        for nav_folder in favorites_group.NavigationFolders:
            print(f" - {nav_folder.Folder.Name}")
            favorites.append(nav_folder.Folder)
            
        return favorites
        
    except Exception as e:
        print(f"Error accessing Favorites: {e}")
        return []

def process_emails_in_favorites():
    """
    Process unread emails in all folders marked as 'Favorites' in Outlook.
    Extract Bill of Lading, PO numbers, or Load IDs (at least 5 digits).
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    favorites_folders = get_favorites_folders(outlook)

    matching_emails = []

    for folder in favorites_folders:
        try:
            items = folder.Items
            items = items.Restrict("[Unread] = True")  # Filter unread emails

            # Print the number of unread emails in the current folder
            print(f"\nChecking folder: '{folder.Name}' - Found {len(items)} unread emails")
            
            for item in items:
                if item.Class == 43:  # MailItem
                    subject = item.Subject or ""
                    body = item.Body or ""

                    # Combine subject and body for processing
                    text = f"{subject} {body}"

                    # Extract numbers
                    numbers = extract_numbers(text)

                    if numbers:
                        matching_emails.append({
                            "Sender": item.SenderName,
                            "Subject": subject,
                            "Numbers": numbers,
                            "EntryID": item.EntryID,  # Store EntryID for future actions like Reply-All
                            "Folder": folder.Name,
                        })

        except Exception as e:
            print(f"Error processing folder {folder.Name}: {e}")
            continue

    return matching_emails


# Main script
if __name__ == "__main__":
    matching_emails = process_emails_in_favorites()

    print("\nMatching Emails:")
    for email in matching_emails:
        print(f"Sender: {email['Sender']}")
        print(f"Subject: {email['Subject']}")
        print(f"Numbers Found: {email['Numbers']}")
        print(f"Folder: {email['Folder']}")
        print("-" * 50)
