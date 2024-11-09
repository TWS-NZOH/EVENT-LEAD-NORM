import os
import csv
import io
import pandas as pd
import traceback
import re
import math
import time
from datetime import datetime, timedelta
from itertools import islice
from fs_norm import FSNormalizer as normalizer

normalizer = normalizer()


# =========================== CSV functions ===========================

# Desktop file retrieval (local file)
def get_desktop_file(filename):
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    file_path = os.path.join(desktop_path, filename)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"The file {filename} was not found on the desktop.")
    with open(file_path, 'rb') as file:
        file_content = file.read()

    return filename, file_content

# Desktop file processing
def normalize_desktop_file(filenames):
    for filename in filenames:
        filename, data = get_desktop_file(filename)
        print(f'PROCESSING {filename}...')
        normalized_df = None

        if 'norm' not in filename.lower():
            # Create a temporary input file
            temp_input = f'temp_input_{filename}'
            with open(temp_input, 'wb') as f:
                f.write(data)
            
            # Create an output filename
            output_string = filename.split('.')[0]
            output_filename = f'{output_string}_normalized.xlsx'

            # Use the normalizer to process the file
            normalizer.normalize_file(temp_input, output_filename)

            # Clean up temporary files
            os.remove(temp_input)

def read_spreadsheet(file_content, file_type):
    if file_type == 'csv':
        df = pd.read_csv(file_content)
    elif file_type == 'xlsx':
        df = pd.read_excel(file_content)
    else:
        raise ValueError(f"Unsupported file format: {file_type}. Please use CSV or XLSX.")
    return df

def split_name(full_name):
    # List of common titles
    titles = ['Mr', 'Mrs', 'Ms', 'Miss', 'Dr', 'Prof']
    
    # Remove any periods and split the name
    name_parts = re.sub(r'\.', '', full_name).split()
    
    # Check if the first part is a title
    if name_parts[0] in titles:
        title = name_parts.pop(0)
    else:
        title = ''
    
    # If there's only one part left, it's treated as the last name
    if len(name_parts) == 1:
        return '', name_parts[0], title
    
    # Otherwise, the last part is the last name, and everything else is the first name
    last_name = name_parts.pop()
    first_name = ' '.join(name_parts)
    
    return first_name, last_name


# Function to parse dates in either mm-dd-yyyy or yyyy-mm-dd format
def parse_date(date_input):
    """
    Parse dates in various formats, including pandas Timestamp objects.
    
    :param date_input: Date in string format (mm-dd-yyyy or yyyy-mm-dd) or pandas Timestamp
    :return: datetime object
    """
    if isinstance(date_input, pd.Timestamp):
        return date_input.to_pydatetime()
    
    if isinstance(date_input, str):
        # Remove any time component if present
        date_str = date_input.split()[0]
        for fmt in ("%m-%d-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                pass
        raise ValueError(f"No valid date format found for {date_input}")
    
    if isinstance(date_input, datetime):
        return date_input
    
    raise ValueError(f"Unsupported date type: {type(date_input)}")

# Function to return numerals from phone numbers; EX '5124592222' instead of '(512) 459-2222'
def format_phone_number(s):
    return ''.join(c for c in s if c.isdigit())

def bulk_check_leads(sf, unprocessed_leads):

    def is_valid_email(email):
        if isinstance(email, str):
            return email.strip() != '' and email.lower() != 'nan'
        return False
    
    # Extract all email addresses
    emails = [lead["account_email"] for lead in unprocessed_leads if is_valid_email(lead["account_email"])]

    # Perform a bulk query to check which accounts exist
    # [ ] ensure that the account type is distro FS || "MBL_Actual_Lead_Source_Name__c": "Fullscript (FS)"
    query = f"""
            SELECT MBL_User_Email__c, Id 
            FROM Account 
            WHERE MBL_User_Email__c IN {tuple(emails)} 
            AND MBL_Account_Channel__c = 'Distributor US'
            AND MBL_Distributor_Accounts__c = 'FullScript (FS)'
            """
    existing_accounts = sf.bulk.Account.query(query)

    # Create a dictionary mapping emails to account IDs
    email_to_account_id = {account['MBL_User_Email__c']: account['Id'] for account in existing_accounts}

    # Separate orders into existing and non-existing accounts
    existing_leads = []
    non_existing_leads = []

    for lead in unprocessed_leads:
        email = lead["account_email"]
        if email in email_to_account_id:
            lead['account_id'] = email_to_account_id[email]
            existing_leads.append(lead)
        else:
            non_existing_leads.append(lead)

    return existing_leads, non_existing_leads

def create_receipt_xlsx(sf_objects, filename):
    flattened_data = []
    
    for sf_object in sf_objects:
        # Extract and flatten customer data
        customer_data = sf_object.get('customer', {})
        flat_customer = {f"{k}": v for k, v in customer_data.items()}
        
        # Extract billing address separately as it's nested
        billing_address = customer_data.get('billing_address', {})
        flat_billing = {f"billing_{k}": v for k, v in billing_address.items()}
        
        shipping_address = customer_data.get('shipping_address', {})
        flat_shipping = {f"shipping_{k}": v for k, v in shipping_address.items()}
        
        # Combine flattened customer and billing data
        flat_customer.update(flat_billing)
        flat_customer.update(flat_shipping)
        
        # Extract order lines
        order_lines = sf_object.get('order_lines', [])
        
        # Extract other top-level keys
        other_data = {k: v for k, v in sf_object.items() if k not in ['customer', 'order_lines']}
        
        # Create a row for each order line
        for line in order_lines:
            row_data = {}
            row_data.update(flat_customer)
            row_data.update(other_data)
            row_data.update(line)
            flattened_data.append(row_data)
    
    # Create DataFrame
    df = pd.DataFrame(flattened_data)
    
    # Save to Excel
    df.to_excel(filename, index=False)
    print(f"[i] {filename} SAVED")

    

