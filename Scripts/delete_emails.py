import sys
import os
import logging

import win32com.client
from datetime import datetime


sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from Supporting_Documents.credentials import MYEMAIL


DELETE_FOLDERS = [
    "Inbox/Weather Updates",
    "Inbox/Coverage",
    "Inbox/BluePrism",
    "Inbox/CFA 2.0",
    "Inbox/CFA Canada",
    "Inbox/CFA Hawaii",
    "Inbox/CFA Contingency",
    "Inbox/CFA Hormel",
    "Inbox/CFA Hubs",
    "Inbox/CFA PR",
    "Inbox/CFA FA",
    "Inbox/CFA MB",
    "Inbox/CFA McLane",
    "Inbox/CFA Perishables",
    "Inbox/CFA QCD",
    "Deleted Items",
]


def delete_app_emails_from_folder(folder_path):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.Folders.Item(MYEMAIL)

    # Handle wildcard pattern
    if "*" in folder_path:
        parent_path = folder_path.replace("/*", "")
        folders = parent_path.split("/")
        current_folder = inbox

        # Navigate to parent folder
        for folder in folders:
            current_folder = current_folder.Folders.Item(folder)

        # Empty all subfolders
        for subfolder in current_folder.Folders:
            while subfolder.Items.Count > 0:
                # Delete each item individually
                subfolder.Items.Item(1).Delete()
    else:
        # Original logic for specific folders
        folders = folder_path.split("/")
        current_folder = inbox
        for folder in folders:
            current_folder = current_folder.Folders.Item(folder)

        # Delete items one by one instead of using Clear()
        while current_folder.Items.Count > 0:
            current_folder.Items.Item(1).Delete()


def execute_app_deletes():
    for folder in DELETE_FOLDERS:
        log_message(f"Folder: {folder}")
        delete_app_emails_from_folder(folder)


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


def setup_logging():
    """
    Sets up logging configuration with a new timestamp-based log file.
    Creates a new log file each time it's called.

    Returns:
        None
    """
    # Create logs directory if it doesn't exist
    log_dir = (
        "C:\\Users\\Zachary Anderson\\Workspace\\ReportProcess\\Scripts\\logs\\deletes"
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
    setup_logging()
    execute_app_deletes()
    log_message("Done deleting")
