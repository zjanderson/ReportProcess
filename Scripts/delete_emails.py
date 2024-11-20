import win32com.client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time


DELETE_FOLDERS = [
    "Inbox/Weather Updates", 
    "Inbox/Fresh Beef", 
    "Inbox/Coverage", 
    "Inbox/BluePrism",
    "Inbox/National Accts/Chik Fil A/*",
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
            
        # Empty all subfolders
        for subfolder in current_folder.Folders:
            subfolder.Items.Clear()  # This is equivalent to "Empty Folder"
    else:
        # Original logic for specific folders
        folders = folder_path.split('/')
        current_folder = inbox
        for folder in folders:
            current_folder = current_folder.Folders.Item(folder)
        
        current_folder.Items.Clear()  # Empty the folder

def execute_app_deletes():
    for folder in DELETE_FOLDERS:
        print(folder)
        delete_app_emails_from_folder(folder)


def delete_web_emails_from_folder(driver, wait, folder_path):
    try:
        # Handle nested folders
        folders = folder_path.split('/')
        
        # Navigate to the parent folder (everything before *)
        for folder in folders:
            if folder == '*':
                # If we hit *, we need to process all subfolders
                return delete_all_subfolders(driver, wait)
            
            folder_xpath = f"//div[contains(@class, 'treeNodeContent')]//span[text()='{folder}']"
            folder_element = wait.until(EC.element_to_be_clickable((By.XPATH, folder_xpath)))
            folder_element.click()
            time.sleep(1)

        # Process current folder if no wildcard
        delete_current_folder(driver, wait, folder_path)

    except TimeoutException as e:
        print(f"Timeout waiting for element in folder {folder_path}: {e}")
    except Exception as e:
        print(f"Error occurred in folder {folder_path}: {e}")

def delete_current_folder(driver, wait, folder_path):
    """Helper function to delete emails in the current folder using Empty folder"""
    try:
        # Right click on the current folder to open context menu
        folder_xpath = f"//div[contains(@class, 'treeNodeContent')]//span[text()='{folder_path.split('/')[-1]}']"
        folder_element = wait.until(EC.element_to_be_clickable((By.XPATH, folder_xpath)))
        
        # Use ActionChains for right click
        from selenium.webdriver.common.action_chains import ActionChains
        actions = ActionChains(driver)
        actions.context_click(folder_element).perform()
        
        # Click "Empty folder" in the context menu
        empty_folder_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//span[text()='Empty folder']")))
        empty_folder_button.click()
        
        # Confirm the deletion in the popup
        confirm_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@data-automation-id='confirmButton']")))
        confirm_button.click()
        
        time.sleep(2)  # Wait for deletion to complete

    except TimeoutException:
        print(f"Could not empty folder {folder_path}")
    except Exception as e:
        print(f"Error emptying folder {folder_path}: {e}")
def delete_all_subfolders(driver, wait):
    """Helper function to recursively delete emails from all subfolders"""
    try:
        # Find all subfolders in current view
        subfolder_xpath = "//div[contains(@class, 'treeNodeContent')]//span"
        subfolders = driver.find_elements(By.XPATH, subfolder_xpath)
        
        if not subfolders:
            print("No subfolders found")
            return

        # Store folder names as clicking will refresh the DOM
        folder_names = [folder.text for folder in subfolders]
        
        for folder_name in folder_names:
            try:
                # Click the subfolder
                folder_xpath = f"//div[contains(@class, 'treeNodeContent')]//span[text()='{folder_name}']"
                folder_element = wait.until(EC.element_to_be_clickable((By.XPATH, folder_xpath)))
                folder_element.click()
                time.sleep(1)

                # Delete emails in this subfolder
                delete_current_folder(driver, wait, folder_name)
                
                # Optional: recursively check for nested subfolders
                delete_all_subfolders(driver, wait)
                
            except Exception as e:
                print(f"Error processing subfolder {folder_name}: {e}")
                continue

    except Exception as e:
        print(f"Error in delete_all_subfolders: {e}")

def execute_web_deletes():
    driver = webdriver.Edge()
    wait = WebDriverWait(driver, 20)  # 20 second timeout
    
    try:
        driver.get("https://outlook.office.com/")
        driver.maximize_window()
        
        # Wait for email interface to load (after login)
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ms-FocusZone")))

        for folder in DELETE_FOLDERS:
            print(folder)
            delete_web_emails_from_folder(driver,wait, folder)
            
    except Exception as e:
        print(f"Error occurred during execution: {e}")


if __name__ == "__main__":
    execute_app_deletes()
    time.sleep(60)
    print("WEB DELETES")
    execute_web_deletes()
