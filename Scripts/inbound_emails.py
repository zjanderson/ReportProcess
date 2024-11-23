import win32com.client
import re
import spacy

# Load spaCy model
nlp = spacy.load("en_core_web_sm")

def extract_numbers(text):
    """
    Extract sequences of at least 5 numerical digits from the text using spaCy.
    """
    numbers = []
    doc = nlp(text)
    for token in doc:
        # Check if the token is a digit and has at least 5 numerical digits
        if token.like_num and len(token.text) >= 5:
            numbers.append(token.text)
    return numbers

def get_favorites_folders(outlook):
    """
    Get folders marked as 'Favorites' in Outlook.
    """
    namespace = outlook.GetNamespace("MAPI")
    favorites = []
    for folder in namespace.Folders:
        try:
            # Check if a folder is marked as a 'Favorite'
            for subfolder in folder.Folders:
                if subfolder.Favorites:
                    favorites.append(subfolder)
        except AttributeError:
            continue
    return favorites

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
            
            for item in items:
                if item.Class == 43:  # MailItem
                    subject = item.Subject or ""
                    body = item.Body or ""

                    # Combine subject and body for processing
                    text = f"{subject} {body}"

                    # Extract numbers using spaCy
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
