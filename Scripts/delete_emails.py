import win32com.client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time


DELETE_FOLDERS = [
    "Inbox/Weather Updates", 
    "Inbox/Fresh Beef", 
    "Inbox/Coverage", 
    "Inbox/BluePrism",
    "Inbox/National Accts/Chik Fil A/*"
    "Deleted Items", 
    ]

def delete_app_emails_from_folder(folder_path):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.Folders.Item("zanderson@armada.net")
    
    # Handle wildcard pattern
    if '*' in folder_path:
        parent_path = folder_path.replace('/*', '')
        folders = parent_path.split('/')
        current_folder = inbox
        
        # Navigate to parent folder
        for folder in folders:
            current_folder = current_folder.Folders.Item(folder)
            
        # Delete from all subfolders
        for subfolder in current_folder.Folders:
            for item in list(subfolder.Items):
                item.Delete()
    else:
        # Original logic for specific folders
        folders = folder_path.split('/')
        current_folder = inbox
        for folder in folders:
            current_folder = current_folder.Folders.Item(folder)
        
        for item in list(current_folder.Items):
            item.Delete()

def execute_app_deletes():
    for folder in DELETE_FOLDERS:
        print(folder)
        delete_app_emails_from_folder(folder)


def delete_web_emails_from_folder(folder_path):
    driver = webdriver.Edge()
    wait = WebDriverWait(driver, 20)  # 20 second timeout

    try:
        driver.get("https://outlook.office.com/")
        driver.maximize_window()
        
        # Wait for email interface to load (after login)
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ms-FocusZone")))

        # Handle nested folders
        folders = folder_path.split('/')
        for folder in folders:
            # More robust folder selection
            folder_xpath = f"//div[contains(@class, 'treeNodeContent')]//span[text()='{folder}']"
            folder_element = wait.until(EC.element_to_be_clickable((By.XPATH, folder_xpath)))
            folder_element.click()

        # Wait for and select all emails
        select_all = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//div[@role='checkbox' and contains(@class, 'checkBox')]")))
        select_all.click()

        # Wait for and click delete
        delete_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@name='Delete']")))
        delete_button.click()

    except TimeoutException as e:
        print(f"Timeout waiting for element: {e}")
    except Exception as e:
        print(f"Error occurred: {e}")
    finally:
        driver.quit()

def execute_web_deletes():
    for folder in DELETE_FOLDERS:
        print(folder)
        delete_web_emails_from_folder(folder)

if __name__ == "__main__":
    # execute_app_deletes()
    # time.sleep(60)
    print("WEB DELETES")
    execute_web_deletes()
