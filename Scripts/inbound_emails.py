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
from selenium.webdriver.common.action_chains import ActionChains

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from Supporting_Documents.credentials import USERNAME, PASSWORD, MYEMAIL

try:
    nltk.download('punkt')
except Exception as e:
    print(f"Error downloading NLTK data: {e}")


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
    # "Darden/DDL", 
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
    "Fresh Beef"
    ]

def extract_numbers(email):
    """
    Extract sequences of at least 5 numerical digits from the email.
    """
    text = email.Subject + email.Body
    numbers = re.findall(r'\b\d{5,}\b', text)
    return set(numbers)

def find_unread_emails(folder_name, inbox):
    try:
        current_folder = inbox.Folders.Item(folder_name)
        
        if current_folder:
            emails = current_folder.Items
            # emails.Sort("[ReceivedTime]", True)
            unread_emails = [email for email in emails if email.UnRead]
            
            if unread_emails:
                print(f"Found {len(unread_emails)} unread emails in {folder_name}")
            else:
                print(f"No unread emails found in {folder_name}")
        return unread_emails
    
    except Exception as folder_error:
        print(f"Error accessing folder '{folder_name}': {folder_error}")

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

def process_email_thread(email):
    numbers = extract_numbers(email)
    print(f"Found {numbers} to search for")

    driver = webdriver.Edge()
    wait = WebDriverWait(driver, 20)

    login_to_tms(driver, wait)
    navigate_to_loads(driver)
    time.sleep(1)

    for number in numbers:
        navigate_to_loads(driver)
        search_in_tms(number, driver)

def process_emails_in_specified_folders():
    """
    Process unread emails that appear to request information in all selected Outlook folders.
    Extract Bill of Lading, PO numbers, or Load IDs (at least 5 digits).
    """
    matching_emails = []

    inbox = access_inbox()



    for folder_name in ALL_FOLDERS:
        unread = find_unread_emails(folder_name, inbox)
        for email in unread:
            process_email_thread(email)
        
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
        driver.get("https://armada.mercurygate.net/MercuryGate/login/spLogin.jsp?")

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

def search_in_tms(number, driver):
    print(f"Searching for number {number} on TMS...")

    actions = ActionChains(driver)
    for _ in range(5):
        actions.send_keys(Keys.TAB).perform()
        time.sleep(.5)   

    actions.send_keys(number)
    actions.send_keys(Keys.RETURN)
    actions.perform()


if __name__ == "__main__":
    numbers = process_emails_in_specified_folders()

    # Use Selenium to navigate and search for numbers
    # driver = webdriver.Edge()
    # wait = WebDriverWait(driver, 20)

    # login_to_tms(driver, wait)
    # navigate_to_loads(driver)
    # search_in_tms("73597774", wait)


    # driver.quit()

