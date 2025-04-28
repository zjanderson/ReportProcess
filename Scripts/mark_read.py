import sys
import os

import win32com.client
from datetime import time, datetime

import logging


ALL_FOLDERS = [
    "IB Hub Dallas",
    "IB Hub East Point",
    "IB Hub Greencastle",
    "IB Hub Romeoville",
    "MCD Toys",
    "MCD East",
    "MCD South",
    "MCD Central",
    "MCD West",
    "MCD Supply",
    "Zaxby's",
    "Bojangles",
    "Stakeholders",
    "Supply Caddy",
    "BBI",
    "CFA Canada",
    "CFA Hawaii",
    "CFA Contingency",
    "CFA Hormel",
    "CFA Hubs",
    "CFA PR",
    "CFA 2.0",
    "CFA FA",
    "CFA MB",
    "CFA McLane",
    "CFA Perishables",
    "CFA QCD",
    "Darden",
    "Darden/DDL",
    "Darden/DDL Maines",
    "Darden/DDL McLane",
    "Dominoes",
    "Panda Express",
    "Panda Produce",
    "Panera",
    "Panera Chips",
    "Panera PandaEx GFS",
    "Panera PandaEx SYGMA",
    "QA",
    "Fresh Beef",
    "Weather Updates",
]


def access_inbox():
    try:
        # Create single Outlook instance outside the loop
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)

        return inbox

    except Exception as e:
        log_message(f"Critical error in Outlook connection: {e}", level="error")

        return e


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


def mark_emails_in_folder_read(folder, military_time):
    try:
        today = datetime.now().date()
        cutoff_time = datetime.combine(today, time(hour=military_time))
        message = f"Cutoff time: {cutoff_time}"
        log_message(message)

        # Get all unread items in the folder
        items = folder.Items.Restrict("[Unread] = True")
        message = f"Found {len(items)} unread emails in {folder.Name}.. marking ones before {military_time}:00 as read"

        count = 0
        for item in items:
            received_time = item.ReceivedTime
            received_datetime = datetime.combine(
                received_time.date(), received_time.time()
            )
            if received_datetime < cutoff_time:
                item.UnRead = False
                item.Save()
                count += 1

        log_message(f"Marked {count} emails as read in folder: {folder.Name}")
        return count

    except Exception as e:
        log_message(f"Error processing folder {folder.Name}: {e}", "error")


def process_folders(time_hours):
    try:
        inbox = access_inbox()
        if isinstance(inbox, Exception):
            log_message(f"Failed to access inbox: {inbox}", level="error")
            return

        for folder_name in ALL_FOLDERS:
            try:
                log_message(f"Attempting to process folder: {folder_name}")
                folder = inbox.Folders.Item(folder_name)
                process_single_folder(folder, time_hours)
            except Exception as e:
                log_message(
                    f"Error accessing folder '{folder_name}': {str(e)}", level="error"
                )
                continue
    except Exception as e:
        log_message(f"Unexpected error in process_folders: {str(e)}", level="error")


def process_single_folder(folder, time_hours, count=10):
    if count > 0:
        new_count = mark_emails_in_folder_read(folder, time_hours)
        if new_count > 0:
            process_single_folder(folder, time_hours, new_count)


def setup_logging():
    """
    Sets up logging configuration with a new timestamp-based log file.
    Creates a new log file each time it's called.

    Returns:
        None
    """
    # Create logs directory if it doesn't exist
    log_dir = "C:\\Users\\Zachary Anderson\\Workspace\\ReportProcess\\Scripts\\logs\\mark_read"
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
    if len(sys.argv) != 2:
        log_message("Usage: python mark_read.py <military_time>", level="error")
        sys.exit(1)

    try:
        military_time = int(sys.argv[1])
        if 0 <= military_time <= 23:
            process_folders(military_time)
            log_message("End file")

        else:
            log_message("Error: Please enter a valid hour between 0 and 23", "error")
            sys.exit(1)

    except ValueError:
        log_message("Error: Please enter a valid integer", "error")
        sys.exit(1)
