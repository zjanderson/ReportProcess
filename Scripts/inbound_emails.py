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
import sys
import os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from Supporting_Documents.credentials import USERNAME, PASSWORD, MYEMAIL

try:
    nltk.download('punkt')
except Exception as e:
    print(f"Error downloading NLTK data: {e}")


ALL_FOLDERS = [
    "Inbox/IB Hub/Dallas", 
    "Inbox/IB Hub/East Point", 
    # "Inbox/IB Hub/Greencastle", 
    # "Inbox/IB Hub/Romeoville", 
    # "Inbox/McDonald's/Toys", 
    # "Inbox/McDonald's/MCD East", 
    # "Inbox/McDonald's/MCD South", 
    # "Inbox/McDonald's/MCD Central", 
    # "Inbox/McDonald's/MCD West", 
    # "Inbox/McDonald's/MCD Supply", 
    # "Inbox/National Accts/Zaxby's", 
    # "Inbox/National Accts/Bojangles", 
    # "Inbox/National Accts/Stakeholders", 
    # "Inbox/National Accts/Supply Caddy", 
    # "Inbox/National Accts/BBI", 
    # "Inbox/National Accts/CHik Fil A/2.0 CFA", 
    # "Inbox/National Accts/CHik Fil A/CFA Canada", 
    # "Inbox/National Accts/CHik Fil A/CFA Hawaii", 
    # "Inbox/National Accts/CHik Fil A/CFA Contingency", 
    # "Inbox/National Accts/CHik Fil A/CFA Hormel", 
    # "Inbox/National Accts/CHik Fil A/CFA Hubs", 
    # "Inbox/National Accts/CHik Fil A/CFA PR", 
    # "Inbox/National Accts/CHik Fil A/FA CFA", 
    # "Inbox/National Accts/CHik Fil A/MB CFA", 
    # "Inbox/National Accts/CHik Fil A/McLane CFA", 
    # "Inbox/National Accts/CHik Fil A/Perishables", 
    # "Inbox/National Accts/CHik Fil A/QCD", 
    # "Inbox/National Accts/Darden", 
    # "Inbox/National Accts/Darden/DDL", 
    # "Inbox/National Accts/Darden/DDL Maines", 
    # "Inbox/National Accts/Darden/DDL McLane", 
    # "Inbox/National Accts/Dominoes", 
    # "Inbox/National Accts/Panda Express", 
    # "Inbox/National Accts/Panera", 
    # "Inbox/National Accts/Panera/Panera Chips", 
    # "Inbox/National Accts/Panera/Panera PandaEx GFS", 
    # "Inbox/National Accts/Panera/Panera PandaEx SYGMA", 
    # "Inbox/QA", 
    ]

def extract_numbers(text):
    """
    Extract sequences of at least 5 numerical digits from the text.
    """
    numbers = re.findall(r'\b\d{5,}\b', text)
    return set(numbers)

def get_folders_to_process(outlook):
    compiled_folders = []
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.Folders.Item(MYEMAIL)
    current_folder = inbox 
    for folder in ALL_FOLDERS:
        compiled_folders.append(folder.Folders)


def process_emails_in_favorites():
    """
    Process unread emails that appear to request information in all folders marked as 'Favorites' in Outlook.
    Extract Bill of Lading, PO numbers, or Load IDs (at least 5 digits).
    """
    outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    favorites_folders = get_folders_to_process(outlook)

    matching_emails = []

    for folder in favorites_folders:
        try:
            emails = folder.Items
            emails = emails.Restrict("[Unread] = True")  # Filter unread emails

            # Print the number of unread emails in the current folder
            print(f"\nChecking folder: '{folder}' - Found {len(emails)} unread emails")
            
            for email in emails:
                if email.Class == 43:  # Mailitem
                    subject = email.Subject or ""
                    body = email.Body or ""

                    # Combine subject and body for processing
                    text = f"{subject} {body}"

                    # Extract numbers
                    numbers = extract_numbers(text)

                    if numbers:
                        matching_emails.append({
                            "Sender": email.SenderName,
                            "Subject": subject,
                            "Numbers": numbers,
                            "EntryID": email.EntryID,  # Store EntryID for future actions like Reply-All
                            "Folder": folder,
                        })

        except Exception as e:
            print(f"Error processing folder {folder}: {e}")
            continue

    return matching_emails

def click_button_by_XPATH(driver, element_xpath):

    wait = WebDriverWait(driver, 10)
    target_element = wait.until(
        EC.element_to_be_clickable((By.XPATH, element_xpath))
    )

    target_element.click()

def login_to_tms(driver, wait):
    """
    Log in to TMS MercuryGate
    """
    try:
        # need to input hardcoded un and pw fields
        username_field = wait.until(EC.presence_of_element_located((By.ID, "UserId")))
        username_field.send_keys(USERNAME)
        
        # # Find password field and enter credentials
        password_field = driver.find_element(By.ID, "Password")
        password_field.send_keys(PASSWORD)

        # Click the Sign In button
        click_button_by_XPATH(driver, '//input[@value="    Sign In    "]')
        
        print("Successfully logged into MercuryGate")


    except Exception as e:
        print(f"Login failed: {e}")
        raise


def navigate_to_loads(driver):

    try:
        # Click the Loads button
        click_button_by_XPATH(driver, '/html/body/table/tbody/tr[2]/td/div[5]/span')
        
        print("Successfully navigated to Loads page")
    
    except Exception as e:
        print(f"Navigation to Loads page failed: {e}")
        raise

def search_in_tms(matching_emails, wait):
    for email in matching_emails:
        for number in email["Numbers"]:
            print(f"Searching for number {number} on TMS...")

            # Example: Perform search
            search_box = wait.until(EC.element_to_be_clickable(By.XPATH, '/html/body/form/table/tbody/tr/td[2]/input[1]'))
            search_box.clear()
            search_box.send_keys(number)
            search_box.send_keys(Keys.RETURN)
            time.sleep(5)  # Wait for results to load


def search_in_second_service(driver, wait, matching_emails):
    """
    Search for numbers in the second cloud service.
    """
    driver.get("https://cloudservice2.com")  # Replace with the second service's URL
    print("Navigating to the second cloud service...")

    for email in matching_emails:
        for number in email["Numbers"]:
            print(f"Searching for number {number} on Cloud Service 2...")
            search_box = wait.until(EC.element_to_be_clickable((By.NAME, "search")))  # Replace NAME with the actual search box locator
            search_box.clear()
            search_box.send_keys(number)
            search_box.send_keys(Keys.RETURN)
            time.sleep(5)  # Wait for results to load


# def navigate_and_search(matching_emails):
    # """
    # Navigate to TMS and 4Kites and search using extracted numbers.
    # """
    # # Set up Selenium WebDriver (Edge, Chrome, or Firefox)
    # driver = webdriver.Edge()  # Replace with webdriver.Chrome() or webdriver.Firefox() as needed
    # wait = WebDriverWait(driver, 20)  # Adjust the timeout as needed

    # try:
    #     # First Cloud Service
    #     driver.get("https://armada.mercurygate.net/MercuryGate/login/mgLogin.jsp?inline=true")
    #     print("Navigating to TMS MercuryGate...")

    #     # Log in 
    #     try:
    #         # need to input hardcoded un and pw fields
    #         username_field = wait.until(EC.presence_of_element_located((By.ID, "UserId")))
    #         username_field.send_keys(USERNAME)  # TMS test environment un is practice
            
    #         # # Find password field and enter credentials
    #         password_field = driver.find_element(By.ID, "Password")
    #         password_field.send_keys(PASSWORD)  # TMS test environment is Armada1@

    #         # Click the Sign In button
    #         click_button_by_XPATH(driver, '//input[@value="    Sign In    "]')
            
    #         print("Successfully logged into MercuryGate")
            
    #     except Exception as e:
    #         print(f"Login failed: {e}")
    #         raise
        
    #     try:
    #         # Wait for the <div> containing the <span> with text "Loads" to be clickable
    #         click_button_by_XPATH(driver, '//*[@id="__AppFrameBaseTable"]/tbody/tr[2]/td/div[5]/span')

    #     except Exception as e:
    #         print(f"Error: {e}")

    #     for email in matching_emails:
    #         for number in email["Numbers"]:
    #             print(f"Searching for number {number} on TMS...")

    #             # Example: Perform search
    #             search_box = wait.until(EC.element_to_be_clickable((By.ID, "search-input")))  # Replace ID with the actual search box locator
    #             search_box.clear()
    #             search_box.send_keys(number)
    #             search_box.send_keys(Keys.RETURN)
    #             time.sleep(5)  # Wait for results to load

    #     # Second Cloud Service
    #     driver.get("https://cloudservice2.com")  # Replace with the second service's URL
    #     print("Navigating to the second cloud service...")

    #     for email in matching_emails:
    #         for number in email["Numbers"]:
    #             print(f"Searching for number {number} on Cloud Service 2...")

    #             # Example: Perform search
    #             search_box = wait.until(EC.element_to_be_clickable((By.NAME, "search")))  # Replace NAME with the actual search box locator
    #             search_box.clear()
    #             search_box.send_keys(number)
    #             search_box.send_keys(Keys.RETURN)
    #             time.sleep(5)  # Wait for results to load

    # except TimeoutException as e:
    #     print(f"Timeout occurred: {e}")
    # except Exception as e:
    #     print(f"An error occurred: {e}")
    # finally:
    #     driver.quit()  # Ensure the browser closes after execution

# Main script
if __name__ == "__main__":
    matching_emails = process_emails_in_favorites()

    # Use Selenium to navigate and search for numbers
    # driver = webdriver.Edge()
    # wait = WebDriverWait(driver, 20)

    # login_to_tms(driver, wait)
    # navigate_to_loads(driver)
    # search_in_tms(matching_emails, wait)

    # driver.quit()

