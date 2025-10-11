import os
import sys
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple
import pandas as pd
import win32com.client as win32

# Configuration
CONFIG = {
    'deferred_delivery_hours': 6,  # Hours to defer email delivery
    'required_columns': ['Carrier Name', 'Dest Name'],  # Required columns in Excel
    'contact_files': {
        'carriers': '..\\Supporting_Documents\\Afterhours_Contacts.xlsx',
        'ops': '..\\Supporting_Documents\\Ops_Contacts.xlsx'
    }
}

# Email template
EMAIL_TEMPLATE = {
    "subject": "Pickup Updates - {carrier_name}",
    "body": """
        <p>Please confirm if the following loads have picked up:</p>
        
        {html_table_with_styles}
        
        <p>If a load has picked up, please update MercuryGate with in/out times, and provide current location and ETA in this thread.</p>
        """,
}


# Validation function
def validate_excel_columns(df: pd.DataFrame, sheet_name: str) -> bool:
    """Validate that required columns exist in the DataFrame."""
    missing_columns = [col for col in CONFIG['required_columns'] if col not in df.columns]
    if missing_columns:
        print(f"❌ Error: Missing required columns in sheet '{sheet_name}': {missing_columns}")
        print(f"Available columns: {list(df.columns)}")
        return False
    return True


# Import big report file
def parse_report(file_name: str, sheet_name: str) -> Optional[pd.core.groupby.DataFrameGroupBy]:
    """Parse Excel file and group by Carrier Name with validation."""
    try:
        df = pd.read_excel(file_name, sheet_name=sheet_name)
        
        # Validate required columns
        if not validate_excel_columns(df, sheet_name):
            return None
        if df.empty:
            print(f"⚠️ Warning: Sheet '{sheet_name}' is empty. No data to process.")
            return None
        else:
            # Group by 'Carrier Name'
            carriers = df.groupby("Carrier Name")
            return carriers
    except (FileNotFoundError, KeyError) as e:
        print(f"❌ Failure to import sheet {sheet_name} from Report file! Error: {e}")
        return None


# Generate table, eliminate NaN
def prepare_data_for_email(group: pd.DataFrame) -> str:
    """Prepare DataFrame for email by cleaning data and creating HTML table."""
    # Eliminate NaN values from the DataFrame
    group = group.fillna(value="")

    # Create HTML table for the current carrier
    table_styles = """
        <style>
        table, th, td {
          border: 1px solid black;
          border-collapse: collapse;
          padding: 5px;
        }
        </style>
        """
    html_table = group.to_html(index=False)  # Convert to HTML table (without index)

    # Add styling to the table
    html_table_with_styles = table_styles + html_table

    return html_table_with_styles


# Sheet processing function
def get_sheet_name(file_name: str) -> str:
    """Get the first (and only) sheet name from Excel file."""
    xl = pd.ExcelFile(file_name)
    available_sheets = xl.sheet_names
    
    if not available_sheets:
        raise ValueError("❌ No sheets found in the workbook!")
    
    if len(available_sheets) > 1:
        print(f"⚠️ Warning: Multiple sheets found. Using first sheet: '{available_sheets[0]}'")
    
    return available_sheets[0]


def initialize_outlook_and_contacts() -> Tuple[object, Dict[str, str], Dict[str, str]]:
    """Initialize Outlook and load contact mappings."""
    try:
        outlook = win32.Dispatch("outlook.application")
        all_carrier_contacts = get_map_carriers_contacts(CONFIG['contact_files']['carriers'])
        email_group = get_map_email_groups(CONFIG['contact_files']['ops'])
        return outlook, all_carrier_contacts, email_group
    except (Exception, FileNotFoundError) as e:
        print(f"❌ Failed to initialize Outlook or load contact maps. Error: {e}")
        raise


def process_carrier_group(
    outlook: object,
    carrier_name: str,
    group: pd.DataFrame,
    all_carrier_contacts: Dict[str, str],
    email_group: Dict[str, str]
) -> None:
    """Process a single carrier group and create email."""
    dest_names = group["Dest Name"].unique()
    html_table_with_styles = prepare_data_for_email(group)

    recipient = all_carrier_contacts.get(carrier_name)
    if not recipient:
        print(f"⚠️ No contact found for carrier: {carrier_name}")
        return

    recipientCC = ";".join(find_CC_recips(dest_names, email_group))

    try:
        mail = compose_email(
            outlook,
            carrier_name,
            recipient,
            recipientCC,
            html_table_with_styles
        )
        mail.Display()
    except (ValueError, AttributeError) as e:
        print(f"❌ Failed to create email for {carrier_name}. Error: {e}")


# Compose a single email with body, signature, and image
def compose_email(
    outlook,
    carrier_name: str,
    recipient: str,
    recipientCC: str,
    html_table_with_styles: str
) -> object:
    """Compose a single email with deferred delivery and signature."""
    # Get signature and image if any
    signature_html, image_file = get_signature_and_image()

    # Create a new email
    mail = outlook.CreateItem(0)  # 0 = Mail item

    mail.Subject = EMAIL_TEMPLATE["subject"].format(carrier_name=carrier_name)
    mail.to = recipient
    mail.cc = recipientCC

    # Set deferred delivery time (configurable)
    delivery_time = datetime.now() + timedelta(hours=CONFIG['deferred_delivery_hours'])
    mail.DeferredDeliveryTime = delivery_time.strftime("%Y-%m-%d %H:%M")

    if image_file:
        attachment = mail.Attachments.Add(image_file)
        # Set Content ID for the image (to embed it in the HTML body)
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "signature_image"
        )

    # Modify signature to reference the embedded image (if applicable)
    if "signature_image" in signature_html:
        signature_html = signature_html.replace('src="', 'src="cid:signature_image"')

    # Create email body
    email_body = EMAIL_TEMPLATE["body"].format(html_table_with_styles=html_table_with_styles)

    # Set the email body (with the table of data)
    mail.HTMLBody = email_body + signature_html

    return mail


# Helper function: Get signature and image
def get_signature_and_image():
    signature_path = os.path.join(os.getenv("APPDATA"), r"Microsoft\Signatures")

    # Check if the signature directory exists
    if not os.path.exists(signature_path):
        return "", None

    # Find the first HTML signature file
    signature_files = [f for f in os.listdir(signature_path) if f.endswith(".htm")]
    if not signature_files:
        return "", None

    # Read the signature HTML
    signature_file = signature_files[0]
    with open(
        os.path.join(signature_path, signature_file), "r", encoding="latin-1"
    ) as f:
        signature_html = f.read()

    # Locate the subfolder with images (if it exists)
    signature_name = os.path.splitext(signature_file)[0]
    image_folder = os.path.join(signature_path, f"{signature_name}_files")
    if os.path.exists(image_folder):
        image_files = [
            f for f in os.listdir(image_folder) if f.endswith((".png", ".jpg", ".jpeg"))
        ]
        if image_files:
            image_file = os.path.join(image_folder, image_files[0])
            # Replace image path in the signature HTML
            signature_html = signature_html.replace(
                'src="', f'src="file:///{image_file}"'
            )
        else:
            image_file = None
    else:
        image_file = None

    return signature_html, image_file


# Helper function: make a hashmap of carrier names and contacts
def get_map_carriers_contacts(contacts_file: str) -> Dict[str, str]:
    """Create a mapping of carrier names to their contact information."""
    try:
        contacts_df = pd.read_excel(contacts_file)
        map_carriers_contacts = {}

        for _, row in contacts_df.iterrows():
            carrier_name = str(row["Carrier"]).strip()
            contact_info = str(row["AFTERHOUR CONTACTS"]).strip()
            map_carriers_contacts[carrier_name] = contact_info
        return map_carriers_contacts
    except (FileNotFoundError, KeyError) as e:
        print(f"❌ Error loading carrier contacts from {contacts_file}: {e}")
        return {}


# Helper function: make a hashmap of locations and email groups
def get_map_email_groups(ops_contacts: str) -> Dict[str, str]:
    """Create a mapping of destination names to email groups."""
    try:
        egroups_df = pd.read_excel(ops_contacts)
        map_email_groups = {}

        for _, row in egroups_df.iterrows():
            dest_name = str(row["Dest Name"]).strip()
            email_group = str(row["Email Group"]).strip()
            map_email_groups[dest_name] = email_group
        return map_email_groups
    except (FileNotFoundError, KeyError) as e:
        print(f"❌ Error loading ops contacts from {ops_contacts}: {e}")
        return {}


# Future enhancement: Owner-based email groups
# TODO: Implement get_map_owner_groups() when Owner_Contacts.xlsx is created
# This would allow for owner-specific email routing based on the Owner column


def find_CC_recips(destinations: List[str], email_group: Dict[str, str]) -> set:
    """Find CC recipients based on destination locations."""
    CC_field = set()

    for location in destinations:
        email = email_group.get(location)
        if email is not None:
            CC_field.add(email)

    return CC_field


# Main email building function
def build_emails(file_name: str) -> None:
    """Main function to build and display emails from Excel file."""
    try:
        # Get the sheet name (only one sheet expected)
        sheet_name = get_sheet_name(file_name)
        print(f"📧 Processing sheet: '{sheet_name}'")

        # Initialize Outlook and contact maps
        outlook, all_carrier_contacts, email_group = initialize_outlook_and_contacts()

        # Parse the report and get carriers
        carriers = parse_report(file_name, sheet_name)
        if carriers is None:
            return

        # Process each carrier group
        for carrier_name, group in carriers:
            process_carrier_group(
                outlook, carrier_name, group, all_carrier_contacts, email_group
            )

    except (FileNotFoundError, ValueError, AttributeError, KeyError) as e:
        print(f"❌ Failed to build emails. Error: {e}")


if __name__ == "__main__":
    env = sys.argv[1]
    if env == "work":
        build_emails("C:\\Users\\zanderson\\Downloads\\Pick.xlsx")
    elif env == "home":
        build_emails("C:\\Users\\Zachary Anderson\\Downloads\\Pick.xlsx")
