import win32com.client
import re
import nltk
from nltk.tokenize import word_tokenize
from nltk.tokenize import sent_tokenize

try:
    nltk.download('punkt')
except Exception as e:
    print(f"Error downloading NLTK data: {e}")

def is_information_request(text):
    """
    Determine if the text contains a question or information request.
    """
    # Common question words and request phrases
    question_indicators = [
        '?', 'what', 'when', 'where', 'who', 'how', 'why', 'can you', 
        'could you', 'please provide', 'please send', 'need to know',
        'looking for', 'requesting', 'inquiry', 'question'
    ]
    
    text = text.lower()
    sentences = sent_tokenize(text)
    
    for sentence in sentences:
        if any(indicator in sentence.lower() for indicator in question_indicators):
            return True
    return False


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
        # Get the active explorer if Outlook is open
        explorer = outlook.Application.ActiveExplorer()
        if not explorer:
            # If no active explorer, try to get default folders
            raise Exception("No active Outlook window found")
            
        # Access the Navigation Pane
        nav_pane = explorer.NavigationPane
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
        # Fallback: return default folders like Inbox
        try:
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # 6 is olFolderInbox
            print("\nFalling back to Inbox only")
            return [inbox]
        except Exception as e2:
            print(f"Fallback also failed: {e2}")
            return []

def process_emails_in_favorites():
    """
    Process unread emails that appear to request information in all folders marked as 'Favorites' in Outlook.
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

                    # Extract numbers if there's info request
                    if is_information_request(text):
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
