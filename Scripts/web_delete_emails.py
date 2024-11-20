from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
import time

def delete_emails_web_outlook(folder_name):
    # Setup WebDriver for Edge
    driver_path = "path_to_edge_webdriver"  # Update with your Edge WebDriver path
    service = Service(driver_path)
    driver = webdriver.Edge(service=service)

    try:
        # Open Outlook Web (logged in)
        driver.get("https://outlook.office.com/")
        driver.maximize_window()
        time.sleep(5)  # Wait for the page to load completely

        # Navigate to the specified folder
        folder = driver.find_element(By.XPATH, f"//span[text()='{folder_name}']")
        folder.click()
        time.sleep(5)  # Wait for the folder to load

        # Select all emails in the folder
        select_all_checkbox = driver.find_element(By.XPATH, "//div[@aria-label='Select all messages']")
        select_all_checkbox.click()
        time.sleep(2)

        # Click Delete
        delete_button = driver.find_element(By.XPATH, "//button[@aria-label='Delete']")
        delete_button.click()

        print(f"Deleted all emails in the folder '{folder_name}'.")
    except Exception as e:
        print(f"Error occurred: {e}")
    finally:
        time.sleep(5)  # Ensure operations complete before closing
        driver.quit()

# Usage
delete_emails_web_outlook("FolderName")  # Replace "FolderName" with your folder name
