import pandas as pd
import win32com.client as win32
import os
import sys
from datetime import datetime, timedelta

# Email templates
EMAIL_TEMPLATES = {
    "overnight_update": {
        "subject": "Overnight Updates - {carrier_name}",
        "body": """
            <p>Please provide updated location and ETA for the following loads:</p>
            
            {html_table_with_styles}
            
            <p>If a load has picked up and/or delivered, please update MercuryGate with in/out times.</p>
            """
    },
    # Add more templates as needed, for example:
    "hot_loads": {
        "subject": "Priority Loads: Status Update Required - {carrier_name}",
        "body": """
            <p>Please provide status updates for the following high-priority loads:</p>
            
            {html_table_with_styles}
            
            <p>If a load has picked up and/or delivered, please update MercuryGate with in/out times.</p>
            """
    }
}

# Import big report file 
def parse_report(file_name, sheet_name):
    try:
        df = pd.read_excel(file_name, sheet_name=sheet_name)
        # Group by 'Carrier Name'
        carriers = df.groupby("Carrier Name")
        return carriers
    except Exception as e: 
        print(f"Failure to import sheet {sheet_name} from Report file! Error defined as: {e}")
        return None


# Generate table, eliminate NaN
def prepare_data_for_email(group):
    # Eliminate NaN values from the DataFrame
    group = group.fillna(value='')

    #  Create HTML table for the current carrier
    table_styles = """
        <style>
        table, th, td {
          border: 1px solid black;
          border-collapse: collapse;
          padding: 5px;
        }
        </style>
        """
    html_table = group.to_html(index=False)  # Convert  to HTML table (without index)

    # Add styling to the table
    html_table_with_styles = table_styles + html_table

    return html_table_with_styles

# Combine sheets 1 and 2 if needed
def combine_sheets(file_name):
    try:
        # Read first two sheets
        df1 = pd.read_excel(file_name, sheet_name=0)  # First sheet
        df2 = pd.read_excel(file_name, sheet_name=1)  # Second sheet
        
        # Combine the dataframes
        combined_df = pd.concat([df1, df2], ignore_index=True)
        
        # Create ExcelWriter object
        with pd.ExcelWriter(file_name, mode='a', if_sheet_exists='replace') as writer:
            # Write combined data to a new sheet
            combined_df.to_excel(writer, sheet_name='Combined', index=False)
            
            # Copy remaining sheets (3 and 4) as is
            df3 = pd.read_excel(file_name, sheet_name=2)
            df4 = pd.read_excel(file_name, sheet_name=3)
            df3.to_excel(writer, sheet_name='Sheet3', index=False)
            df4.to_excel(writer, sheet_name='Sheet4', index=False)
            
        return True
        
    except Exception as e:
        print(f"Failed to combine sheets. Error: {e}")
        return False



# Compose a single email with body, signature, and image
def compose_email(outlook, carrier_name, recipient, recipientCC, html_table_with_styles, template_key="overnight_update"):
    # Get signature and image if any
    signature_html, image_file = get_signature_and_image()

    # Create a new email
    mail = outlook.CreateItem(0)  # 0 = Mail item
    template = EMAIL_TEMPLATES.get(template_key)
    if not template:
        raise ValueError(f"Invalid template key: {template_key} not found")
    
    mail.Subject = template["subject"].format(carrier_name=carrier_name)
    mail.to = recipient
    mail.cc = recipientCC

    # Set deferred delivery time to 3 hours from now
    from datetime import datetime, timedelta
    delivery_time = datetime.now() + timedelta(hours=3)
    mail.DeferredDeliveryTime = delivery_time.strftime("%Y-%m-%d %H:%M")

    if image_file:
        attachment = mail.Attachments.Add(image_file)
        # Set Content ID for the image (to embed it in the HTML body)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "signature_image")

    # Modify signature to reference the embedded image (if applicable)
    if "signature_image" in signature_html:
        signature_html = signature_html.replace("src=\"", "src=\"cid:signature_image\"")

    # Create email body
    email_body = template["body"].format(html_table_with_styles=html_table_with_styles)
    
    # Set the email body (with the table of data)
    mail.HTMLBody = email_body + signature_html
    
    return mail


# Helper function: Get signature and image
def get_signature_and_image():
    signature_path = os.path.join(os.getenv('APPDATA'), r"Microsoft\Signatures")

    # Check if the signature directory exists
    if not os.path.exists(signature_path):
        return "", None

    # Find the first HTML signature file
    signature_files = [f for f in os.listdir(signature_path) if f.endswith('.htm')]
    if not signature_files:
        return "", None

    # Read the signature HTML
    signature_file = signature_files[0]
    with open(os.path.join(signature_path, signature_file), 'r', encoding='latin-1') as f:
        signature_html = f.read()

    # Locate the subfolder with images (if it exists)
    signature_name = os.path.splitext(signature_file)[0]
    image_folder = os.path.join(signature_path, f"{signature_name}_files")
    if os.path.exists(image_folder):
        image_files = [f for f in os.listdir(image_folder) if f.endswith(('.png', '.jpg', '.jpeg'))]
        if image_files:
            image_file = os.path.join(image_folder, image_files[0])
            # Replace image path in the signature HTML
            signature_html = signature_html.replace('src="', f'src="file:///{image_file}"')
        else:
            image_file = None
    else:
        image_file = None

    return signature_html, image_file

# Helper function: make a hashmap of carrier names and contacts 
def get_map_carriers_contacts(contacts_file):
    contacts_df = pd.read_excel(contacts_file)

    map_carriers_contacts = {}

    for row_number, row in contacts_df.iterrows():
        carrier_name = str(row['Carrier']).strip()
        contact_info = str(row['AFTERHOUR CONTACTS']).strip()

        map_carriers_contacts[carrier_name] = contact_info 
    return map_carriers_contacts

# Helper function: make a hashmap of locations and email groups
def get_map_email_groups(ops_contacts):
    # Read all sheets from the Excel file
    all_sheets = pd.read_excel(ops_contacts, sheet_name=None)
    
    map_email_groups = {}
    
    # Process each sheet
    for sheet_name, sheet_df in all_sheets.items():
        # Only process sheets with "DC" in their titles or the "IBHub" sheet
        if ("DC" in sheet_name or sheet_name == "IBHub") and 'Dest Name' in sheet_df.columns and 'Email Group' in sheet_df.columns:
            for row_number, row in sheet_df.iterrows():
                dest_name = str(row['Dest Name']).strip()
                email_group = str(row['Email Group']).strip()
                
                # Only add if both values are not empty
                if dest_name and email_group and dest_name != 'nan' and email_group != 'nan':
                    map_email_groups[dest_name] = email_group
    
    return map_email_groups

# Helper function: make a hashmap of owners and email groups
def get_map_owner_groups(ops_contacts):
    # Read all sheets from the Excel file
    all_sheets = pd.read_excel(ops_contacts, sheet_name=None)
    
    map_owner_groups = {}
    
    # Process each sheet
    for sheet_name, sheet_df in all_sheets.items():
        # Only process sheets with "DC" in their titles or the "IBHub" sheet
        if ("DC" in sheet_name or sheet_name == "IBHub") and 'Owner' in sheet_df.columns and 'Email Group' in sheet_df.columns:
            for row_number, row in sheet_df.iterrows():
                owner = str(row['Owner']).strip()
                email_group = str(row['Email Group']).strip()
                
                # Only add if both values are not empty
                if owner and email_group and owner != 'nan' and email_group != 'nan':
                    map_owner_groups[owner] = email_group
    
    return map_owner_groups

# Helper function - finding CC field of email groups for McD and CFA - check 'Owner' column for .contains MCD or Chik-fil-a, then reference destinations, otherwise new hashmap for Owner 
# and corresponding email group
def find_CC_recips(destinations, email_group, owner_group=None, owner=None):
    CC_field = set()

    # Add email groups based on destinations
    for location in destinations:
        email = email_group.get(location)
        if email is not None:
            CC_field.add(email)
    
    # Add email groups based on owner if provided
    if owner_group and owner:
        owner_email = owner_group.get(owner)
        if owner_email is not None:
            CC_field.add(owner_email)

    return CC_field

# Build and Display emails
def build_emails(file_name):
    try:
        # Get available sheets from the Excel file
        xl = pd.ExcelFile(file_name)
        available_sheets = xl.sheet_names
        sheet_count = len(available_sheets)
        
        if sheet_count == 0:
            print("No sheets found in the workbook!")
            return
        elif sheet_count == 4:
            print("Found 4 sheets. Combining sheets 1 and 2...")
            if combine_sheets(file_name):
                # Refresh Excel file handle after modification
                xl = pd.ExcelFile(file_name)
                available_sheets = ['Combined', 'Sheet3', 'Sheet4']
                sheet_count = 3
            else:
                print("Failed to combine sheets. Exiting.")
                return
        elif sheet_count > 4:
            print("Warning: More than 4 sheets found. Only processing the first 4.")
            available_sheets = available_sheets[:4]
            sheet_count = 4
            
        print(f"Processing {sheet_count} sheets")
        
        # Initialize Outlook and contact maps
        outlook = win32.Dispatch('outlook.application')
        all_carrier_contacts = get_map_carriers_contacts("..\\Supporting_Documents\\Afterhours_Contacts.xlsx")
        email_group = get_map_email_groups("..\\Supporting_Documents\\Ops_Contacts.xlsx")
        
        # Process sheets in reverse order
        for i in range(sheet_count - 1, -1, -1):
            sheet_name = available_sheets[i]
            # Determine template based on sheet count and position
            template_key = 'overnight_update' if (sheet_count == 3 and i == 0) else 'hot_loads'
            
            print(f"Processing sheet {sheet_name} with template: {template_key}...")
            
            carriers = parse_report(file_name, sheet_name)
            if carriers is None:
                continue
                
            for carrier_name, group in carriers:
                dest_names = group['Dest Name'].unique()
                html_table_with_styles = prepare_data_for_email(group)
                
                recipient = all_carrier_contacts.get(carrier_name)
                if not recipient:
                    print(f"No contact found for carrier: {carrier_name}")
                    continue
                    
                recipientCC = ";".join(find_CC_recips(dest_names, email_group))
                
                try:
                    mail = compose_email(
                        outlook, 
                        carrier_name, 
                        recipient, 
                        recipientCC, 
                        html_table_with_styles,
                        template_key=template_key
                    )
                    mail.Display()
                except Exception as e:
                    print(f"Failed to create email for {carrier_name} in {sheet_name}. Error: {e}")
                    
    except Exception as e:
        print(f"Failed to initialize Outlook or load contact maps. Error: {e}")

# Build overnight update emails for each carrier
def build_overnight_update_emails(carriers, ops_contacts):
    # Get email groups from the new spreadsheet
    email_groups = get_map_email_groups(ops_contacts)
    owner_groups = get_map_owner_groups(ops_contacts)
    
    emails = []
    
    for carrier_name, carrier_data in carriers:
        # Prepare data for email
        html_table = prepare_data_for_email(carrier_data)
        
        # Get unique destinations for this carrier
        destinations = carrier_data['Dest Name'].unique().tolist()
        
        # Get owner from the first row (assuming all rows have the same owner)
        owner = None
        if 'Owner' in carrier_data.columns and not carrier_data['Owner'].empty:
            owner = str(carrier_data['Owner'].iloc[0]).strip()
        
        # Find CC recipients based on destinations and owner
        CC_field = find_CC_recips(destinations, email_groups, owner_groups, owner)
        
        # Create email
        email = {
            "to": f"{carrier_name}@example.com",  # Replace with actual carrier email
            "cc": list(CC_field),
            "subject": EMAIL_TEMPLATES["overnight_update"]["subject"].format(carrier_name=carrier_name),
            "body": EMAIL_TEMPLATES["overnight_update"]["body"].format(html_table_with_styles=html_table)
        }
        
        emails.append(email)
    
    return emails

if __name__ == "__main__":
    env = sys.argv[1]
    if env == "work":
        build_emails("C:\\Users\\zanderson\\Downloads\\Report.xlsx")
    elif env == "home":
        build_emails("C:\\Users\\Zachary Anderson\\Downloads\\Report.xlsx")