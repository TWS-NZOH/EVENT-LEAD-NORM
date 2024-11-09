import pandas as pd
import traceback
import openpyxl
import math
import csv
import re
import os

class FSNormalizer:
    def __init__(self):
        
        self.path_to_desktop = '/Users/tws/Desktop/'
        self.processing_subfolder = 'LEAD-NORM/'
        
        # profession normalization  
        self.professions_substitutions = {
            'Pharmacist' : 'RPH - Registered Pharmacist',
            'Nutritional Counselor' : 'NC - Nutrition Counselor',
            'Chiropractor' : 'DC - Doctor of Chiropractic',
            'Naturopathic Doctor' : 'ND - Naturopathic Doctor',
            'Health Coach' : 'HC - Health Coach',
            'Naturopathic Doctor degree' : 'ND - Naturopathic Doctor',
            'Medical Doctor' : 'MD - Doctor of Medicine',
            'Dietitian' : 'RDN - Registered Dietitian Nutritionist',
            'Osteopathic Physician' : 'MD - Doctor of Medicine',
            'Acupuncturist' : 'LAc - Licensed Acupuncturist',
            'Nurse Practitioner' : 'NP - Nurse Practitioner/Family Nurse Practitioner',
            'Herbalist' : 'Other',
            'Physical/Occupational Therapist' : 'OT - Occupational Therapist',
            'Registered Nurse' : 'RN - Registered Nurse',
            'Licensed Nutritionist' : 'CN - Certified Nutritionist',
            'Other' : 'Other',
            'Nutritionist' : 'CN - Certified Nutritionist',
            'Other Licensed Healthcare Provider' : 'Other',
            'Physician Assistant' : 'PA - Physician Assistant',
            'Holistic Healthcare Provider' : 'Other',
            'Registered Midwife' : 'Other',
            'Homeopath' : 'Other',
            'Fitness Professional' : 'PT - Personal Trainers',
            'Practitioner of Oriental Medicine' : 'Other',
            'Massage Therapist' : 'Other',
            'Counselor' : 'LPC - Licensed Professional Counselor',
            'Optometrist' : 'OD - Doctor of Optometry',
            'Veterinary Medicine' : 'DVM - Doctor of Veterinary Medicine',
            'Dentist' : 'DMD - Doctor of Dental Medicine',
            'Eastern Medicine' : 'Other',
            'Healthcare degree' : 'Other',
            'Doctor degree' : 'MD - Doctor of Medicine',
            'Nurse degree' : 'RN - Registered Nurse',
            'Registered Dietitian' : 'RDN - Registered Dietitian Nutritionist',
            'Licensed Nurse Practitioner' : 'NP - Nurse Practitioner/Family Nurse Practitioner',
            'STU' : 'Student',
            'Licensed Osteopathic Physician' : 'MD - Doctor of Medicine',
            'Licensed Medical Doctor' : 'MD - Doctor of Medicine',
            'Certified Nutritionist' : 'CN - Certified Nutritionist',
            'NUTR' : 'CN - Certified Nutritionist',
            'Licensed Naturopathic Doctor' : 'ND - Naturopathic Doctor',
            'DC' : 'DC - Doctor of Chiropractic',
            'LAC' : 'LAc - Licensed Acupuncturist',
            'Licensed Chiropractor' : 'DC - Doctor of Chiropractic',
            'Licensed Physician Assistant' : 'PA - Physician Assistant',
            'Registered Pharmacist' : 'RPH - Registered Pharmacist',
            'Licensed Acupuncturist' : 'LAc - Licensed Acupuncturist',
            'Licensed Physical/Occupational Therapist' : 'OT - Occupational Therapist',
            'Medical / Osteopathic Doctor' : 'MD - Doctor of Medicine',
            'NP' : 'NP - Nurse Practitioner/Family Nurse Practitioner',
            'Naturopathic Doctor (CNME)' : 'ND - Naturopathic Doctor',
            'ND' : 'ND - Naturopathic Doctor',
            'HC' : 'HC - Health Coach',
            'RD' : 'RDN - Registered Dietitian Nutritionist',
            'LYNUT' : 'NC - Nutrition Counselor',
            'MD' : 'MD - Doctor of Medicine',
            'LYNAT' : 'Other',
            'DDS' : 'DMD - Doctor of Dental Medicine',
            'Nurse' : 'RN - Registered Nurse',
            'LicNUR' : 'RN - Registered Nurse',
            'LicOTH' : 'Other',
            'LICPT' : 'OT - Occupational Therapist',
            'PA' : 'PA - Physician Assistant',
            'Licensed Dentist' : 'DMD - Doctor of Dental Medicine',
            'RPH' : 'RPH - Registered Pharmacist',
            'Licensed Counselor' : 'LPC - Licensed Professional Counselor',
            'DVM' : 'DVM - Doctor of Veterinary Medicine',
            'LYOTH' : 'Other',
            'Licensed Massage Therapist' : 'Other',
            'DO' : 'MD - Doctor of Medicine',
            'LicCL' : 'LPC - Licensed Professional Counselor',
            'Licensed Optometrist' : 'OD - Doctor of Optometry',
            'LicMT' : 'Other',
            'Licensed Practitioner of Oriental Medicine' : 'Other',
            'LYBW' : 'Other',
            'HOUSE' : 'Other',
            'Licensed Veterinary Practitioner' : 'DVM - Doctor of Veterinary Medicine',
            'Healthcare Professional' : 'Other',
            'AHG' : 'Other',
            'DPM' : 'Other',
            'QAP' : 'LAc - Licensed Acupuncturist',
            'HS' : 'Other',
            'FITPR' : 'PT - Personal Trainers'     
            }
        
        self.us_states = {
            'AL': 'Alabama',
            'AK': 'Alaska',
            'AZ': 'Arizona',
            'AR': 'Arkansas',
            'CA': 'California',
            'CO': 'Colorado',
            'CT': 'Connecticut',
            'DE': 'Delaware',
            'FL': 'Florida',
            'GA': 'Georgia',
            'HI': 'Hawaii',
            'ID': 'Idaho',
            'IL': 'Illinois',
            'IN': 'Indiana',
            'IA': 'Iowa',
            'KS': 'Kansas',
            'KY': 'Kentucky',
            'LA': 'Louisiana',
            'ME': 'Maine',
            'MD': 'Maryland',
            'MA': 'Massachusetts',
            'MI': 'Michigan',
            'MN': 'Minnesota',
            'MS': 'Mississippi',
            'MO': 'Missouri',
            'MT': 'Montana',
            'NE': 'Nebraska',
            'NV': 'Nevada',
            'NH': 'New Hampshire',
            'NJ': 'New Jersey',
            'NM': 'New Mexico',
            'NY': 'New York',
            'NC': 'North Carolina',
            'ND': 'North Dakota',
            'OH': 'Ohio',
            'OK': 'Oklahoma',
            'OR': 'Oregon',
            'PA': 'Pennsylvania',
            'RI': 'Rhode Island',
            'SC': 'South Carolina',
            'SD': 'South Dakota',
            'TN': 'Tennessee',
            'TX': 'Texas',
            'UT': 'Utah',
            'VT': 'Vermont',
            'VA': 'Virginia',
            'WA': 'Washington',
            'WV': 'West Virginia',
            'WI': 'Wisconsin',
            'WY': 'Wyoming'
            }

    def add_column(self, headers, data, column_name, after_column=None):
        if after_column is None:
            headers.append(column_name)
        else:
            for i, header in enumerate(headers):
                if header == after_column:
                    headers.insert(i+1, column_name)
                    break
        for row in data:
            row[column_name] = ''

    def remove_column(self, headers, data, column_name):
        for i, header in enumerate(headers):
            if header == column_name:
                del headers[i]
                break
        for row in data:
            del row[column_name]

    def rename_column(self, headers, data, column_name, new_column_name):
        for i, header in enumerate(headers):
            if header == column_name:
                headers[i] = new_column_name
                break
        for row in data:
            row[new_column_name] = row[column_name]
            del row[column_name]

    def vlookup(self, filename, lookup_column, lookup_value, return_column):
        desktop_filename = f'{self.path_to_desktop}{filename}'
        with open(desktop_filename, 'r', encoding='utf-8') as vf:
            reader = csv.DictReader(vf)
            for row in reader:
                if row[lookup_column] == lookup_value:
                    return row[return_column]

    def normalize_zip(self, zip_code):
        if pd.isna(zip_code):
            return zip_code
        zip_code = str(zip_code)
        if "-" in zip_code:
            zip_code = zip_code.split("-")[0]
        return zip_code[:5].zfill(5)

    def split_name(self, full_name):
        if pd.isna(full_name):
            return pd.Series(['', ''])
    
        list_of_salutations = ["Mrs.","Mrs ","Mr. ","Mr ","Miss ","Dr. ","Dr ","Ms. ","Ms "]
        
        for salutation in list_of_salutations:
            if full_name.startswith(salutation):
                full_name = full_name.replace(salutation, "").strip()
                break
        
        full_name = re.sub(r',.*', '', full_name)
        parts = full_name.split(maxsplit=1)
        return pd.Series([parts[0], parts[-1] if len(parts) > 1 else ''])
    # ================

    def ensure_dir_exists(self, file_path):
        # Ensure directory exists
        directory = os.path.dirname(file_path)
        if directory and not os.path.exists(directory):
            os.makedirs(directory)

    def normalize_data(self, df):
     
        # Apply normalization logic
        # Address handling
        df['Street'] = df['Street'].astype(str).str.strip() if len(str(df['Street'])) > 3 else ''
        df['Street2'] = df['Street2'].astype(str).str.strip() if len(str(df['Street'])) > 3 else ''
        df['Street'] = df.apply(lambda row: f"{row['Street']}, {row['Street2']}" if 'nan' not in str(row['Street2']).lower() else row['Street'] if 'nan' not in str(row['Street']).lower() else '', axis=1)

        # ZIP code normalization
        df['PostalCode'] = df['PostalCode'].astype(str).apply(self.normalize_zip) if 'nan' not in str(df['PostalCode']).lower() else ''

        # Account modality mapping
        df['MBL_Profession__c'] = df['MBL_Profession__c'].astype(str).map(self.professions_substitutions) if len(str(df['MBL_Profession__c'])) > 3 else ''

        # Rename columns
        df = df.rename(columns={'First Name': 'first_name', 'Last Name': 'last_name', 'Company': 'account_name', 'Email': 'account_email', 
                                 'Street': 'address', 'Street2': 'address_2', 'Phone': 'account_phone', 'CountryCode': 'country_code',
                                 'StateCode': 'state_code', 'PostalCode': 'zip_postal', 'LeadSource': 'lead_source', 'City': 'city'})

        # Remove unwanted columns
        columns_to_remove = ['address_2']
        
        df = df.drop(columns=columns_to_remove, errors='ignore')

        return df

    def normalize_file(self, input_file, output_file):

        df = None

        try:
            # Read the file into a DataFrame based on its extension
            _, file_extension = os.path.splitext(input_file)
            if file_extension.lower() == '.xlsx':
                df = pd.read_excel(input_file)
            elif file_extension.lower() == '.csv':
                df = pd.read_csv(input_file)
            else:
                raise ValueError(f"Unsupported file type: {file_extension}")

            # Ensure the output directory exists
            file_path = f'{self.path_to_desktop}{self.processing_subfolder}{output_file}'
            self.ensure_dir_exists(file_path)

            # Normalize the data
            df = self.normalize_data(df)

            # Write the output
            df.to_excel(file_path, index=False)

            print(f"[+] SUCCESS! Output written to: {file_path}")

        except Exception as e:
            print(f"An error occurred while processing the file: {str(e)}")
            print("FS Norm Traceback:")
            print(traceback.format_exc())
        
        return df