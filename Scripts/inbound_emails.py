import win32com.client
import re
import nltk
from nltk.tokenize import word_tokenize
from nltk.tokenize import sent_tokenize
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time

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

def click_button_by_XPATH(driver, element_xpath):

    wait = WebDriverWait(driver, 10)
    target_element = wait.until(
        EC.element_to_be_clickable((By.XPATH, element_xpath))
    )

    target_element.click()


def navigate_and_search(matching_emails):
    """
    Navigate to TMS and 4Kites and search using extracted numbers.
    """
    # Set up Selenium WebDriver (Edge, Chrome, or Firefox)
    driver = webdriver.Edge()  # Replace with webdriver.Chrome() or webdriver.Firefox() as needed
    wait = WebDriverWait(driver, 20)  # Adjust the timeout as needed

    try:
        # First Cloud Service
        driver.get("https://armada.mercurygate.net/MercuryGate/login/mgLogin.jsp?inline=true")
        print("Navigating to TMS MercuryGate...")

        # Log in 
        try:
            # need to input hardcoded un and pw fields
            username_field = wait.until(EC.presence_of_element_located((By.ID, "UserId")))
            username_field.send_keys("zachary.anderson")  # TMS test environment un is practice
            
            # # Find password field and enter credentials
            password_field = driver.find_element(By.ID, "Password")
            password_field.send_keys("di&6Rwt2#f7PB6")  # TMS test environment is Armada1@

            # Click the Sign In button
            click_button_by_XPATH(driver, '//input[@value="    Sign In    "]')
            # wait = WebDriverWait(driver, 10)
            # sign_in_button = wait.until(
            #     EC.element_to_be_clickable((By.XPATH, '//input[@value="    Sign In    "]'))
            #     )

            # # Click the button
            # sign_in_button.click()
            
            # Wait for successful login (adjust the selector based on a element that appears after login)
            wait.until(EC.presence_of_element_located((By.ID, "dashboard")))
            print("Successfully logged into MercuryGate")
            
        except Exception as e:
            print(f"Login failed: {e}")
            raise
        
        try:
            # Wait for the <div> containing the <span> with text "Loads" to be clickable
            click_button_by_XPATH(driver, '//*[@id="__AppFrameBaseTable"]/tbody/tr[2]/td/div[5]/span')

        except Exception as e:
            print(f"Error: {e}")

        finally:
            driver.quit()

        # # Example: Wait for login or dashboard page to load
        # wait.until(EC.presence_of_element_located((By.ID, "dashboard")))

        for email in matching_emails:
            for number in email["Numbers"]:
                print(f"Searching for number {number} on TMS...")

                # Example: Perform search
                search_box = wait.until(EC.element_to_be_clickable((By.ID, "search-input")))  # Replace ID with the actual search box locator
                search_box.clear()
                search_box.send_keys(number)
                search_box.send_keys(Keys.RETURN)
                time.sleep(5)  # Wait for results to load

        # Second Cloud Service
        driver.get("https://cloudservice2.com")  # Replace with the second service's URL
        print("Navigating to the second cloud service...")

        for email in matching_emails:
            for number in email["Numbers"]:
                print(f"Searching for number {number} on Cloud Service 2...")

                # Example: Perform search
                search_box = wait.until(EC.element_to_be_clickable((By.NAME, "search")))  # Replace NAME with the actual search box locator
                search_box.clear()
                search_box.send_keys(number)
                search_box.send_keys(Keys.RETURN)
                time.sleep(5)  # Wait for results to load

    except TimeoutException as e:
        print(f"Timeout occurred: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        driver.quit()  # Ensure the browser closes after execution


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

    # Use Selenium to navigate and search for numbers
    navigate_and_search(matching_emails)

