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
from PIL import Image
import pytesseract
import io
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' # Adjust path if needed

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
    try:
        print(f"Searching for number {number} on TMS...")

        actions = ActionChains(driver)
        for _ in range(5):
            actions.send_keys(Keys.TAB).perform()
            time.sleep(.5)   
        actions.send_keys(number)
        actions.send_keys(Keys.RETURN)
        actions.perform()

        print("Search completed successfully")
        
    except Exception as e:
        print(f"Error during search: {e}")  

def get_shipper_details_tms(driver, wait):
    """
    Extract shipper details from results table using screenshot and OCR
    """
    try:
        # Take screenshot of the entire page
        print("Taking screenshot of page...")
        screenshot = driver.get_screenshot_as_png()
        image = Image.open(io.BytesIO(screenshot))
        
        # Use pytesseract to extract text from the image
        print("Extracting text from screenshot...")
        text = pytesseract.image_to_string(image)
        
        # Parse the extracted text for relevant information
        shipper_details = {}
        lines = text.split('\n')
        with open('lines.py', 'w', encoding='utf-8') as f:
            f.write('lines = ' + repr(lines))
        print("Lines saved to lines.py")
        print(lines)
    except Exception as e:
        print(f"Error processing screenshot: {str(e)}")
        return None



def get_destination_details(driver, wait):
    """
    Extract destination details from the consignee table
    """
    try:
        # Wait for the destination cells with DetailBodyTableRowEven class
        consignee_section = wait.until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div:contains('Consignee')"))
        )
        
        consignee_table = consignee_section.find_element(By.CSS_SELECTOR, "table")
        destination_cells = consignee_table.find_elements(By.CLASS_NAME, "DetailBodyTableRowEven")
        
        destination_data = {}
        
        # Process each cell to extract information
        for cell in destination_cells:
            content = cell.text.strip()
            
            # Skip empty cells
            if not content:
                continue
                
            # Map the content to appropriate dictionary keys
            if 'Contact :' in content:
                destination_data['contact'] = content.replace('Contact :', '').strip()
            if 'Phone :' in content:
                destination_data['phone'] = content.replace('Phone :', '').strip()
            if 'Email :' in content:
                destination_data['email'] = content.replace('Email :', '').strip()
            if 'Location Comments :' in content:
                destination_data['comments'] = content.replace('Location Comments :', '').strip()
            # if 'US' in content:  # Likely the city/state/zip line
            #     destination_data['location'] = content
            # if 'Pkwy' in content or 'Street' in content or 'Road' in content:  # Likely the street address
            #     destination_data['street'] = content
            # if content and 'company' not in destination_data:  # First non-empty cell is usually company name
            #     destination_data['company'] = content
        
        print("Extracted destination details:", destination_data)
        print(destination_data)
        return destination_data
        
    except TimeoutException:
        print("Could not find destination details")
        return None


def dump_page_info(driver, identifier=""):
    try:
        # Get the specific table element
        table = driver.find_element(By.ID, "__AppFrameBaseTable")
        table_html = table.get_attribute('outerHTML')
        
        # Create a timestamp for unique filename
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        filename = f"page_dump_{identifier}_{timestamp}.txt"
        
        # Write table content to file
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(table_html)
                
        print(f"Table content dumped to {filename}")
        return filename
        
    except Exception as e:
        print(f"Error dumping table info: {e}")
        return None

if __name__ == "__main__":
    # numbers = process_emails_in_specified_folders()
    edge_options = webdriver.EdgeOptions()
    edge_options.set_capability('ms:loggingPrefs', {'performance': 'ALL'})

    # Use Selenium to navigate and search for numbers
    driver = webdriver.Edge(options=edge_options)
    wait = WebDriverWait(driver, 20)

    login_to_tms(driver, wait)
    navigate_to_loads(driver)
    navigate_to_loads(driver)

    search_in_tms("504198867", driver)
    time.sleep(2)
    # dump_page_info(driver, "504198867")

    shipper_details = get_shipper_details_tms(driver, wait)
    # print("\nshipper details:", shipper_details)

    input("\nPress Enter to close the browser.")


    driver.quit()


    

