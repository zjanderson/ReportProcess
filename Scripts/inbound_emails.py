"""Script to process inbound emails and extract load information from TMS.

This module handles email processing, load number extraction, and automated
responses with contact details retrieved from MercuryGate TMS system.
"""

import re
import time
import sys
import os
import io
import logging
from datetime import datetime

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

from Supporting_Documents.credentials import USERNAME, PASSWORD

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

try:
    nltk.download("punkt")
except (OSError, IOError, Exception) as e:
    print(f"Error downloading NLTK data: {e}", "error")

ALL_FOLDERS = [
    # "IB Hub Dallas",
    # "IB Hub East Point",
    # "IB Hub Greencastle",
    # "IB Hub Romeoville",
    # "MCD Toys",
    "MCD East",
    # "MCD South",
    "MCD Central",
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
    "Panda Express",
    "Panda Produce",
    "Panera",
    "Panera Chips",
    "Panera PandaEx GFS",
    "Panera PandaEx SYGMA",
    # "QA",
    # "Fresh Beef",
]


def access_inbox():
    """
    Establishes a connection to the Outlook application and

    Returns:
    Inbox: the Outlook inbox folder object if successful
    Exception: The error object if failure
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)

        return inbox

    except (AttributeError, OSError, RuntimeError) as e:
        log_message(f"Critical error in Outlook connection: {e}", "error")
        return e


def click_button_by_xpath(driver, element_xpath):
    """
    Uses Selenium driver to click a button given the relevant xpath
    """

    wait = WebDriverWait(driver, 10)
    target_element = wait.until(EC.element_to_be_clickable((By.XPATH, element_xpath)))

    target_element.click()


def compose_body(extracted_number, shipper_details, consignee_details):
    """
    Composes an HTML-formatted email body with contact information for shipper and consignee
      with the particular number from the email
    """
    if (
        not shipper_details["emails"]
        and not shipper_details["phone_numbers"]
        and not consignee_details["emails"]
        and not consignee_details["phone_numbers"]
    ):
        body = (
            f"<pre> {extracted_number}: Load found but no details could be extracted. "
            f"Maybe check manually <br></pre>"
        )

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
    """
    Creates a reply all email and saves it in drafts
    """
    try:
        reply = email.ReplyAll()
        if reply is None:
            log_message("Error: Could not create reply")
            return False
        reply.HTMLBody = body + reply.HTMLBody

        # Display the email (this returns an Inspector object)
        # reply.Display()

        # Save the draft
        reply.Save()

        log_message("Response email composed and saved as draft")
        return True
    except (AttributeError, RuntimeError) as e:
        log_message(f"Error composing response email: {e}", "error")
        return False


def execute_all_email_actions():
    """
    Main execution function that process all unread emails.
    Coordinates the entire process from email retrieval to response generation.
    """
    setup_logging()
    unread_emails = extract_all_unread_emails()
    for email in unread_emails:
        body = extract_all_details_for_thread(email)
        compose_response_email(email, body)
        mark_as_read(email)

    log_message("End file")


def handle_screenshot(driver):
    """
    Takes a screenshot of the current page and processes it for OCR.

    Args:
        driver: The Selenium WebDriver instance

    Returns:
        PIL.Image: Processed image ready for OCR
    """
    screenshot = driver.get_screenshot_as_png()
    image = Image.open(io.BytesIO(screenshot))

    image = image.convert("L")  # Convert to grayscale
    # Increase contrast
    image = image.point(lambda x: 0 if x < 128 else 255, "1")

    return image


def decide_state(driver):
    """
    Determines the state of the TMS page by analyzing OCR text.

    Args:
        driver: The Selenium WebDriver instance

    Returns:
        str: State of the page - "loads found", "no loads found", or "unknown"
    """
    image = handle_screenshot(driver)
    custom_config = (
        r'--oem 3 --psm 6 -c tessedit_char_whitelist="ABCDEFGHIJKLMNOPQRSTUVWXYZ"'
        r'abcdefghijklmnopqrstuvwxyz0123456789@.,_-:/ ()"'
    )
    text = pytesseract.image_to_string(image, config=custom_config)
    if "noloadsfound." in text.lower():
        print("DEFINITELY NO LOADS")
        return "no loads found"
    elif "phone" in text.lower():
        print("DEFINITELY LOADS")
        return "loads found"
    print("UNKNOWN")
    return "unknown"


def handle_loads_found(driver, number):
    """
    Handles the case when loads are found in TMS.

    Args:
        driver: The Selenium WebDriver instance
        number: The load number being processed

    Returns:
        str: Formatted body text with contact details
    """
    shipper_details = get_contact_details_tms(driver, "shipper")
    log_message(f"Shipper details: {shipper_details}")
    consignee_details = get_contact_details_tms(driver, "consignee")
    log_message(f"Consignee details: {consignee_details}")
    number_body = compose_body(number, shipper_details, consignee_details)
    return number_body


def process_single_number(number, driver):
    """
    Process a single number through TMS, handling all the necessary steps.
    Returns a tuple of (success, result) where:
    - success is a boolean indicating if the processing was successful
    - result is either the body text if successful, or None if failed
    """
    try:
        navigate_to_loads(driver)
        navigate_to_loads(driver)
        navigate_to_loads(driver)

        if not search_in_tms(number, driver):
            return False, None

        time.sleep(2)
        state = decide_state(driver)
        if state == "loads found":
            return True, handle_loads_found(driver, number)
        elif state == "no loads found":
            return True, f"Definitely no loads for {number} \n\n"
        else:
            log_message(f"Unknown state for number {number}")
            return False, None

    except (TimeoutError, RuntimeError, AttributeError) as e:
        log_message(f"Error processing number {number}: {str(e)}", "error")
        return False, None


def process_numbers(driver, numbers, processed_numbers=None, attempt=1, max_retries=3):
    """
    Process a list of numbers through TMS, handling retries recursively.

    Args:
        driver: The Selenium WebDriver instance
        numbers: List or deque of numbers to process
        processed_numbers: Set of numbers that have been successfully processed
        attempt: Current attempt number for retries
        max_retries: Maximum number of retry attempts allowed

    Returns:
        str: The accumulated body text from successfully processed numbers
    """
    if processed_numbers is None:
        processed_numbers = set()

    total_body = ""
    retry_numbers = []

    for number in numbers:
        if number in processed_numbers:
            continue

        success, result = process_single_number(number, driver)
        if success:
            total_body += result
            processed_numbers.add(number)
        else:
            retry_numbers.append(number)

    # If we have numbers to retry and haven't exceeded max retries, process them recursively
    if retry_numbers and attempt < max_retries:
        log_message(f"Retrying {len(retry_numbers)} numbers (attempt {attempt + 1})")
        total_body += process_numbers(
            driver, retry_numbers, processed_numbers, attempt + 1, max_retries
        )
    elif retry_numbers:
        log_message(f"Max retries reached for {len(retry_numbers)} numbers", "warning")

    return total_body


def extract_all_details_for_thread(email):
    """
    Extracts load numbers from an email thread, retrieves contact details for each load from TMS,
    combines all contact info into a single return string
    """
    numbers = extract_numbers(email)
    log_message(f"Found {numbers} to search for")
    if len(numbers) == 0:
        return "No numbers found"

    try:
        log_message("Initializing Edge options...")
        edge_options = webdriver.EdgeOptions()
        edge_options.add_argument("--no-sandbox")
        edge_options.add_argument("--disable-dev-shm-usage")
        edge_options.add_argument("--disable-gpu")
        edge_options.add_argument("--disable-extensions")
        edge_options.add_argument("--disable-software-rasterizer")
        edge_options.add_argument("--disable-notifications")
        # edge_options.add_argument("--headless=new")
        edge_options.add_argument("--start-maximized")
        edge_options.set_capability("ms:loggingPrefs", {"performance": "ALL"})

        log_message("Attempting to create WebDriver instance...")
        try:
            driver = webdriver.Edge(options=edge_options)
            log_message("WebDriver instance created successfully")
        except Exception as e:
            log_message(f"Failed to create WebDriver instance: {str(e)}", "error")
            raise

        try:
            log_message("Setting up WebDriverWait...")
            wait = WebDriverWait(driver, 20)
            log_message("WebDriverWait configured successfully")
        except Exception as e:
            log_message(f"Failed to set up WebDriverWait: {str(e)}", "error")
            driver.quit()
            raise

        try:
            log_message("Attempting to login to TMS...")
            login_to_tms(driver, wait)
            log_message("Successfully logged into TMS")
        except Exception as e:
            log_message(f"Failed to login to TMS: {str(e)}", "error")
            driver.quit()
            raise

        # Process all numbers (including retries)
        total_body = process_numbers(driver, numbers)

        driver.quit()
        return total_body

    except (TimeoutError, RuntimeError, AttributeError, OSError) as e:
        log_message(f"Error in extract_all_details_for_thread: {str(e)}", "error")
        if "driver" in locals():
            try:
                driver.quit()
            except (RuntimeError, AttributeError):
                pass
        return f"Error in search. Look up numbers {numbers}. \n Error message: {str(e)}"


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
    """
    Uses regex to find emails in a given text and returns all matches in a list
    """
    email_pattern = r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"
    emails = re.findall(email_pattern, text)
    return emails


def find_phone_numbers(text):
    """
    Uses regex to find phone numbers in a given text and returns all matches in a list
    """
    phone_pattern = (
        r"\b(?:\+?1[-.]?)?\s*(?:\([0-9]{3}\)|[0-9]{3})[-.\s]*[0-9]{3}[-.\s]*[0-9]{4}\b"
    )
    phones = re.findall(phone_pattern, text)
    return phones


def find_unread_emails(folder_name, inbox):
    """
    Finds all unread emails in a specified Outlook folder, returns a list of unread email objects
    """
    try:
        current_folder = inbox.Folders.Item(folder_name)

        if current_folder:
            unread_emails = current_folder.Items.Restrict("[Unread] = True")
            # emails.Sort("[ReceivedTime]", True)

            if unread_emails:
                log_message(
                    f"Found {len(unread_emails)} unread emails in {folder_name}"
                )
            else:
                log_message(f"No unread emails found in {folder_name}")
        return unread_emails

    except (AttributeError, RuntimeError) as folder_error:
        log_message(f"Error accessing folder '{folder_name}': {folder_error}", "error")


def get_contact_details_tms(driver, details_type):
    """
    Extract contact details from results table using screenshot and OCR
    """
    try:
        # Take screenshot of the entire page
        log_message("Taking screenshot of page...")
        image = handle_screenshot(driver)
        width, height = image.size
        # crop image to only show shipper (left third) or consignee (middle third)
        if details_type == "shipper":
            image = image.crop((0, 0, width // 3, height))
        if details_type == "consignee":
            image = image.crop(
                (width // 3, 0, 2 * width // 3, height)
            )  # Crop to middle third of screen

        custom_config = (
            r'--oem 3 --psm 6 -c tessedit_char_whitelist="ABCDEFGHIJKLMNOPQRSTUVWXYZ"'
            r'abcdefghijklmnopqrstuvwxyz0123456789@.,_-:/ ()"'
        )
        text = pytesseract.image_to_string(image, config=custom_config)
        emails = find_emails(text)
        phone_numbers = find_phone_numbers(text)
        contact_detials = {"emails": emails, "phone_numbers": phone_numbers}
        return contact_detials
    except (OSError, IOError, RuntimeError) as e:
        log_message(f"Error processing screenshot: {str(e)}", "error")
        return None


def log_message(message, level="info"):
    """
    Logs a message to both console and file.

    Args:
        message (str): The message to log
        level (str): The logging level ('info', 'error', 'warning', 'debug')
    """
    print(message)  # Print to console
    if level == "error":
        logging.error(message)
    elif level == "warning":
        logging.warning(message)
    elif level == "debug":
        logging.debug(message)
    else:
        logging.info(message)


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
        click_button_by_xpath(driver, '//input[@value="    Sign In    "]')

        log_message("Successfully logged into MercuryGate")

    except (TimeoutError, RuntimeError, AttributeError) as e:
        log_message(f"Login failed: {e}", "error")
        raise


def mark_as_read(email):
    """
    Marks a particular email object as read and saves it
    """
    try:
        email.UnRead = False
        email.Save()
        log_message("Email marked as read")
    except (AttributeError, RuntimeError) as e:
        log_message(f"Error marking email as read: {e}", "error")


def navigate_to_loads(driver):
    """
    Uses Selenium to navigate to the loads tab in Mercury Gate
    """

    try:
        # Click the Loads button
        click_button_by_xpath(driver, "/html/body/table/tbody/tr[2]/td/div[5]/span")

        log_message("Successfully navigated to Loads page")

    except (TimeoutError, RuntimeError, AttributeError) as e:
        log_message(f"Navigation to Loads page failed: {e}", "error")
        raise


def search_in_tms(number, driver):
    """
    Uses Selenium to search TMS Mercury Gate for a particular number.
    Returns True if search was successful, False if it failed and should be retried.
    """
    try:
        log_message(f"Searching for number {number} on TMS...")

        actions = ActionChains(driver)
        for _ in range(5):
            actions.send_keys(Keys.TAB).perform()
            time.sleep(0.5)
        actions.send_keys(number)
        actions.send_keys(Keys.RETURN)
        actions.perform()
        time.sleep(1)

        # Check for alert
        try:
            alert = driver.switch_to.alert
            alert_text = alert.text
            alert.accept()
            log_message(f"Alert encountered: {alert_text}", "warning")
            return False
        except (RuntimeError, AttributeError):
            # No alert found, search was successful
            log_message("Search completed successfully")
            return True

    except (TimeoutError, RuntimeError, AttributeError) as e:
        log_message(f"Error during search: {e}", "error")
        return False


def setup_logging():
    """
    Sets up logging configuration with a new timestamp-based log file.
    Creates a new log file each time it's called.

    Returns:
        None
    """
    # Create logs directory if it doesn't exist
    log_dir = (
        "C:\\Users\\Zachary Anderson\\Workspace\\ReportProcess\\Scripts\\logs\\inbound"
    )
    os.makedirs(log_dir, exist_ok=True)

    # Remove any existing handlers
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)

    # Generate timestamp-based filename
    current_time = datetime.now().strftime("%Y%m%d-%H%M")
    log_file = os.path.join(log_dir, f"log-{current_time}.txt")

    # Set up logging configuration
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    log_message("Logging initialized with new log file")


if __name__ == "__main__":
    execute_all_email_actions()
