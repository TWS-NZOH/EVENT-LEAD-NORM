# Leads from events automation
'''
SOP Steps:

1. Save the file you want to normalize to the desktop
2. Modify the string on line 53 to match the exact formatting and file name of the file you saved to the desktop
3. Do the same in the fs_norm file on line 12
3. Run the script by running the following commands in the terminal: 
    a. Make sure you're in the right folder with:
        cd /Users/KSPA/Desktop/LEAD-NORM/
    b. Run the script with:
        python3 leads_norm.py
4. The script will normalize the file on the desktop and save it to the desktop with the starting name prepended with 'normalized_'
'''

# =========================== IMPORTS ===========================

import sys
import site
sys.path.extend(site.getsitepackages())

import gc # garbage collection for freeing up memory
import pandas as pd 
from fs_norm import FSNormalizer
from leads_norm_functions import *
# ///////// end csv imports ///////////

import sys
import time
from datetime import datetime

print(f'// RUNNING [EVENT LEADS NORMALIZATION V.001] //')

# ========================= INITIALIZATION =========================

# We want to time our operations, so we start our timer
start_time = time.time()
total_timer_start = time.time()

normalizer = FSNormalizer()
all_errors = {}
desktop_file_name_1 = 'event-leads.xlsx'
desktop_file_name_2 = None
desktop_file_names = [desktop_file_name_1, desktop_file_name_2] if desktop_file_name_2 is not None else [desktop_file_name_1]

# Print the action duration and reset our timer
end_time = time.time()
print(f'[+] INITIALIZATION COMPLETE: {round((end_time - start_time), 2)} SECONDS')
start_time = time.time()

# ========================= LEAD NORMALIZATION =========================

# Normalize the desktop file
normalize_desktop_file(desktop_file_names)

date = datetime.today().strftime("%Y-%m-%d")
start_time = time.time()

total_timer_end = time.time()
print(f'================ TOTAL RUN TIME =================')
print(f'{round((total_timer_end - total_timer_start), 2)} SECONDS')
print(f'==================================================')


