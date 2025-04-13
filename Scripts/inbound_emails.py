import re
import time
import sys
import os
import io

import win32com.client
import nltk
import pytesseract

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.action_chains import ActionChains
from PIL import Image

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from Supporting_Documents.credentials import USERNAME, PASSWORD

try:
    nltk.download("punkt")
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


def click_button_by_XPATH(driver, element_xpath):

    wait = WebDriverWait(driver, 10)
    target_element = wait.until(EC.element_to_be_clickable((By.XPATH, element_xpath)))

    target_element.click()


def compose_body(extracted_number, shipper_details, consignee_details):
    if (
        not shipper_details["emails"]
        and not shipper_details["phone_numbers"]
        and not consignee_details["emails"]
        and not consignee_details["phone_numbers"]
    ):
        body = f"<pre> {extracted_number}: no details found <br></pre>"

    else:
        body = f"""
        <pre>
        Number: {extracted_number}

            Shipper details:
                emails: {shipper_details['emails']}
                phones: {shipper_details['phone_numbers']}

            Consignee details:
                emails: {consignee_details['emails']}
                phones: {consignee_details['phone_numbers']}
        </pre>
        """
    return body


def compose_response_email(email, body):
    try:
        reply = email.ReplyAll()
        if reply is None:
            print("Error: Could not create reply")
            return False
        reply.HTMLBody = body + reply.HTMLBody

        # Display the email (this returns an Inspector object)
        reply.Display()

        # Save the draft
        reply.Save()

        print("Response email composed and saved as draft")
        return True
    except Exception as e:
        print(f"Error composing response email: {e}")
        return False


def execute_all_email_actions():
    unread_emails = extract_all_unread_emails()
    for email in unread_emails:
        body = extract_all_details_for_thread(email)
        compose_response_email(email, body)
        mark_as_read(email)


def extract_all_details_for_thread(email):
    numbers = extract_numbers(email)
    total_body = ""
    print(f"Found {numbers} to search for")

    edge_options = webdriver.EdgeOptions()

    edge_options.set_capability("ms:loggingPrefs", {"performance": "ALL"})
    edge_options.add_argument("--headless")

    # Use Selenium to navigate and search for numbers
    driver = webdriver.Edge(options=edge_options)
    wait = WebDriverWait(driver, 20)
    login_to_tms(driver, wait)

    for number in numbers:
        print(number)
        navigate_to_loads(driver)
        navigate_to_loads(driver)
        navigate_to_loads(driver)
        search_in_tms(number, driver)
        time.sleep(2)
        shipper_details = get_contact_details_tms(driver, "shipper")
        print("\nshipper details: ", shipper_details)
        consignee_details = get_contact_details_tms(driver, "consignee")
        print("consignee_details: ", consignee_details)
        number_body = compose_body(number, shipper_details, consignee_details)
        total_body += number_body
    driver.quit()
    return total_body


def extract_all_unread_emails():
    """
    Find all unread emails in the folders specified in ALL_FOLDERS
    Return a list of all those unread emails
    """
    all_unread_emails = []

    inbox = access_inbox()

    for folder_name in ALL_FOLDERS:
        unread = find_unread_emails(folder_name, inbox)
        all_unread_emails.extend(unread)

    return all_unread_emails


def extract_numbers(email):
    """
    Extract sequences of at least 5 numerical digits from the email.
    """
    text = email.Subject + email.Body
    numbers = re.findall(r"\b\d{5,}\b", text)
    return set(numbers)


def find_emails(text):
    email_pattern = r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"
    emails = re.findall(email_pattern, text)
    return emails


def find_phone_numbers(text):
    phone_pattern = (
        r"\b(?:\+?1[-.]?)?\s*(?:\([0-9]{3}\)|[0-9]{3})[-.\s]*[0-9]{3}[-.\s]*[0-9]{4}\b"
    )
    phones = re.findall(phone_pattern, text)
    return phones


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


def get_contact_details_tms(driver, details_type):
    """
    Extract contact details from results table using screenshot and OCR
    """
    try:
        # Take screenshot of the entire page
        print("Taking screenshot of page...")
        screenshot = driver.get_screenshot_as_png()
        image = Image.open(io.BytesIO(screenshot))
        width, height = image.size
        # crop image to only show shipper (left third) or consignee (middle third)
        if details_type == "shipper":
            image = image.crop((0, 0, width // 3, height))
        if details_type == "consignee":
            image = image.crop(
                (width // 3, 0, 2 * width // 3, height)
            )  # Crop to middle third of screen

        image = image.convert("L")  # Convert to grayscale
        image = image.point(lambda x: 0 if x < 128 else 255, "1")  # Increase contrast
        custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789@.,_-:/ ()"'
        text = pytesseract.image_to_string(image, config=custom_config)
        emails = find_emails(text)
        phone_numbers = find_phone_numbers(text)
        contact_detials = {"emails": emails, "phone_numbers": phone_numbers}
        return contact_detials
    except Exception as e:
        print(f"Error processing screenshot: {str(e)}")
        return None


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


def mark_as_read(email):
    try:
        email.UnRead = False
        email.Save()
        print("Email marked as read")
    except Exception as e:
        print(f"Error marking email as read: {e}")


def navigate_to_loads(driver):

    try:
        # Click the Loads button
        click_button_by_XPATH(driver, "/html/body/table/tbody/tr[2]/td/div[5]/span")

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
            time.sleep(0.5)
        actions.send_keys(number)
        actions.send_keys(Keys.RETURN)
        actions.perform()

        print("Search completed successfully")

    except Exception as e:
        print(f"Error during search: {e}")


if __name__ == "__main__":
    execute_all_email_actions()
