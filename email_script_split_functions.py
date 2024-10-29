import pandas as pd
import win32com.client as win32
import os


# Function 1: Import big report file for one email
def parse_report(file_name):
    try:
        excel_file = file_name
        df = pd.read_excel(excel_file)
    
        # Group by 'Carrier Name'
        carriers = df.groupby("Carrier Name")
        return carriers
    except Exception as e: 
        print(f"Failure to import Report file! Error defined as: {e}")


# Function 2: Generate table, eliminate NaN
def prepare_data_for_email(group):
    # Eliminate NaN values from the DataFrame
    group = group.fillna('')

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


# Function 3: Compose a single email with body, signature, and image
def compose_email(outlook, carrier_name, recipient,recipientCC, html_table_with_styles):
    # Get signature and image if any
    signature_html, image_file = get_signature_and_image()

    # Create a new email
    mail = outlook.CreateItem(0)  # 0 = Mail item
    mail.Subject = f"Overnight Updates - {carrier_name}"
    mail.to = recipient
    mail.cc = recipientCC

    # Attach image (if any)
    if image_file:
        attachment = mail.Attachments.Add(image_file)
        # Set Content ID for the image (to embed it in the HTML body)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "signature_image")

    # Modify signature to reference the embedded image (if applicable)
    if "signature_image" in signature_html:
        signature_html = signature_html.replace("src=\"", "src=\"cid:signature_image\"")

    # Create email body
    email_body = f"""
        <p>Please provide updated location and status for the following loads:</p>
        
        {html_table_with_styles}
        
        <p>If a load has picked up and/or delivered, please update MercuryGate with in/out times.</p>
        """
    
    # Set the email body (with the table of data)
    mail.HTMLBody = email_body + signature_html
    
    return mail


# Helper function: Get signature and image
def get_signature_and_image():
    signature_path = os.path.join(os.getenv('APPDATA'), r"Microsoft\Signatures")

    if os.path.exists(signature_path):
        # Find the .htm file (HTML signature)
        signature_files = [f for f in os.listdir(signature_path) if f.endswith('.htm')]
        
        if signature_files:
            signature_name = signature_files[0].split(".htm")[0]  # Get signature name without the extension
            with open(os.path.join(signature_path, signature_files[0]), 'r', encoding='latin-1') as f:
                signature_html = f.read()
            
            # Locate the subfolder where images are stored (subfolder has the same name as the signature)
            image_subfolder = os.path.join(signature_path, signature_name + "_files")
            
            if os.path.exists(image_subfolder):
                image_files = [f for f in os.listdir(image_subfolder) if f.endswith(('.png', '.jpg', '.jpeg'))]
                if image_files:
                    image_file = os.path.join(image_subfolder, image_files[0])
                else:
                    image_file = None
            else:
                image_file = None
        else:
            signature_html = ""
            image_file = None
    else:
        signature_html = ""
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
    egroups_df = pd.read_excel(ops_contacts)

    map_email_groups = {}

    for row_number, row in egroups_df.iterrows():
        dest_name = str(row['Dest Name']).strip()
        email_group = str(row['Email Group']).strip()

        map_email_groups[dest_name] = email_group
    return map_email_groups

# Helper function: make a hashmap of owners and email groups ##MAKE SPREADSHEET OWNER_CONTACTS WITH OWNERS AND CORRESPONDING EMAIL GROUPS, add to build funcion
#def get_map_owner_groups(owner_contacts):
#    egroups_df = pd.read_excel(owner_contacts)

#    map_owner_groups = {}

#    for row_number, row in egroups_df.iterrows():
#        owner = str(row['Owner']).strip()
#        email_group = str(row['Email Group']).strip()

#        map_owner_groups[owner] = email_group
#    return map_owner_groups



# Helper function - finding CC field of email groups for McD and CFA - check 'Owner' column for .contains MCD or Chik-fil-a, then reference destinations, otherwise new hashmap for Owner 
# and corresponding email group

def find_CC_recips(destinations, email_group):

    CC_field = ''

    for location in destinations:
        email = email_group.get(location)
        if email is not None:
            CC_field+=email_group.get(location)
            CC_field+=(';')

    return CC_field


# Function 4: Send emails
def build_emails(file_name):
    # Parse report
    carriers = parse_report(file_name)

    # Initialize Outlook
    outlook = win32.Dispatch('outlook.application')

    # all contacts hashmap
    all_carrier_contacts = get_map_carriers_contacts("C:\\Users\\zanderson\\Documents\\Afterhours_Contacts.xlsx")

    # Email groups hashmap
    email_group = get_map_email_groups("C:\\Users\\zanderson\\Documents\\Ops_Contacts.xlsx")

    # Loop through each unique Carrier and send an email
    for carrier_name, group in carriers:
        dest_name = group.get('Dest Name')
        #print(dest_name)

        # Normalize data and prepare HTML table
        html_table_with_styles = prepare_data_for_email(group)

        recipient = all_carrier_contacts.get(carrier_name)

        recipientCC = find_CC_recips(dest_name, email_group)

        
        # Compose the email
        try:
            mail = compose_email(outlook, carrier_name, recipient, recipientCC, html_table_with_styles)

            # Display the email (use mail.Send() to send directly)
            mail.Display()
        except Exception as e:
            print(f"Failed for {carrier_name}. Error {e}.")



# Run the script
build_emails("C:\\Users\\zanderson\\Downloads\\Report.xlsx")

#get_map_carriers_contacts("C:\\Users\\zanderson\\Documents\\Afterhours_Contacts.xlsx")

#TO DO LIST, ROUGHLY PRIORITIZED

#create Repo for version control, access from both machines to test.

#Data entry of spreadsheet for email groups

#update function to build email with CC mailto fields - partially complete

#add functionality to schedule send of emails at randomized intervals based on either discreet start time or NOW + X hours

#QC to see if CFA/McD/other owners share location info, and how to parse that.

#build function to do normal loads AND function to do priority loads (change title and body of email only)

#update functions to be able to access different sheets of workbook?
