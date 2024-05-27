import os
import logging
import time
from datetime import datetime, timedelta
from tkinter import *
from tkinter import messagebox
import shutil
import glob # For file pattern matching
import subprocess # For file opening
import win32com.client # For detecting if file is open and then closing it
import pandas as pd
import openpyxl

# ---------------------------- BACKEND SETUP ------------------------------- #
# Logging (enhanced for more informative messages)
logging.basicConfig(
    level=logging.DEBUG, 
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Active Directory
# Get the absolute path of the script and use it as active directory
active_directory = os.path.dirname(os.path.abspath(__file__))
os.chdir(active_directory)

# ---------------------------- FILE MANAGEMENT ----------------------------- #
tiktok_ads_path = "Inputs/Thrasio TikTok Shop-Campaign Report(2024-05-12 to 2024-05-18).xlsx"
all_orders_path = "Inputs/All order-2024-05-20-16_03.csv"
affiliate_orders_path = "Inputs/all_20240511210000_20240518205959.csv"
video_analytics_path = "Inputs/Video Performance List_20240520200916.xlsx"
insense_transactions_path = "Inputs/transactions_history (1).xlsx"
company_catalog_path = "Inputs/TikTok Shop US Product Lists.xlsx"
weekly_dashboard_path = "Outputs/TikTok Processor.xlsx"

def create_backup(file_path):
    """Creates a timestamped backup and deletes previous backups of the given file.
    
    Args:
        file_path (string): file path for the file that will be handled.
    """
    try:
        # Backup file creation
        file_dir, file_name = os.path.split(file_path)
        backup_name = os.path.join(file_dir, f"{os.path.splitext(file_name)[0]}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx")
        shutil.copy(file_path, backup_name)

        # Delete previous backups of the same file
        backup_pattern = os.path.join(file_dir, f"{os.path.splitext(file_name)[0]}_*.xlsx")
        for old_backup in glob.glob(backup_pattern):
            if old_backup != backup_name:
                try:
                    os.remove(old_backup)
                    logger.info(f"Deleted previous backup: {old_backup}")
                except Exception as e:
                    logger.warning(f"Failed to delete backup: {old_backup}. Error: {e}")

        logger.info(f"Backup created: {backup_name}")
        logger.info(f"Backup created at: {os.path.abspath(backup_name)}")
    except Exception as e:
        logger.error(f"Backup or deletion failed: {e}")
        messagebox.showerror(title="Error", message=f"An error occurred during backup: {e}")

def open_file(file_path):
    """Open the file using the default application.
    
    Args:
        file_path (string): file path for the file that will be handled.
    """
    try:
        # Open the file with the default associated application
        subprocess.Popen(["start", "", file_path], shell=True)
        logger.info(f"open_file: Opened the file: {file_path}")
    except Exception as e:
        logger.error(f"open_file: An error occurred: {e}")
        messagebox.showerror(title="Error", message=f"An error occurred while opening the file: {e}")
        
def close_open_file(file_path):
    """Check if the specified file is open and close it if necessary.
    
    Args:
        file_path (string): file path for the file that will be handled.
    """
    try:
        # Normalize and get the absolute file path for the operating system, with lower case drive letter
        normalized_file_path = os.path.abspath(os.path.normpath(file_path)).lower()

        # Create an instance of the Excel application
        excel = win32com.client.Dispatch("Excel.Application")

        # Debug: Log the normalized file path
        logger.debug(f"Normalized file path to close: {normalized_file_path}")

        # Extract the filename from the file path
        target_filename = os.path.basename(normalized_file_path)

        # Iterate through the open workbooks
        for workbook in excel.Workbooks:
            workbook_path = os.path.abspath(os.path.normpath(workbook.FullName)).lower()

            # Debug: Log the workbook path
            logger.debug(f"Open workbook path: {workbook_path}")

            # Extract the filename from the workbook path
            workbook_filename = os.path.basename(workbook_path)

            # Compare filenames to handle OneDrive paths and case differences
            if workbook_filename == target_filename:
                workbook.Close(SaveChanges=False)
                logger.info(f"close_open_file: Closed the open file: {normalized_file_path}")
                break
        else:
            logger.info(f"close_open_file: File is not open: {normalized_file_path}")

        # Ensure proper cleanup
        del excel
        time.sleep(1)  # Add a short delay to ensure the Excel process releases the file
    except Exception as e:
        logger.error(f"close_open_file: An error occurred: {e}")
        messagebox.showerror(title="Error", message=f"An error occurred while closing the file: {e}")

# ---------------------------- ERROR CHECKERS ------------------------------ #
def check_write_permission(directory):
    """Check if the user has writer permission to create file backup

    Args:
        directory (string): current directory where the code is being run from

    Returns:
        None: Uses return to stop the code from executing as no backup can be made
    """    
    if os.access(directory, os.W_OK):
        logger.info(f"Write permission exists for: {directory}")
        return True
    else:
        logger.error(f"No write permission for: {directory}")
        return False
    
# ---------------------------- CONSTANTS ------------------------------- #
FONT_NAME = "Calibri"
FONT_SIZE = 11
WHITE = "#fcf7f9"
PLATFORM_FEE = 0.06   # Tiktok's Platform Fee

# ASINs to Process
B07N316V8C = "Angry Orange Pet Odor Eliminator with Citrus Scent for Strong Dog or Cat Pee Smells on Carpet, Furniture & Indoor Outdoor Floors - 24 Fluid Ounces - Puppy Supplies"
B001PU9A9Q = "Nippies Skin Reusable Covers - Sticky Adhesive Silicone Pasties - Reusable Skin Adhesive Covers for Women with Travel Box"
# Dictionary of TTPID that should be processed
ACTIVE_PRODUCT_LIST = {"1729385211989037378": B07N316V8C,
                       "1729386030694895938": B001PU9A9Q
                       } 

# ---------------------------- FINANCIAL METRICS --------------------------- #
sales = None
seller_discount = None
shipping_fee_income = None
shipping_fee_seller_discount = None
shipping_fee_net_income = None
gross_revenue = None
returns = None
net_revenue = None
cogs = None
shipping_cost = None
total_cogs = None
product_sample_cogs = None
product_sample_shipping_cost = None
affiliate_commission = None
insense_joinbrands_flat_fee = None
marketing_affiliate_total = None
tiktok_ads = None
total_marketing_expenses = None
platform_fee = None
net_income_cp = None
net_margin = None
asp_before_disc = None
asp_after_disc = None
total_units = None
total_units_wow_percent = None
ads_impressions = None
video_views = None
product_impressions = None
product_clicks = None
media_ctr = None
media_cvr = None
num_of_samples_sent_insense = None
num_of_samples_sent_tts = None
num_of_content_posted = None
viral_video_greater_than_1mm_vv = None
viral_video_link = None
viral_video_performance = None
best_video_views = None
best_video_link = None
best_video_performance = None
lw_vv_avg = None
event = None
comment = None

# ---------------------------- WEEKLY TASKS -------------------------------- #
def weekly_tasks(ttpid):
    """Runs the weekly tasks as defined in the instructions.
    These are activated via a button in the UI.\n

    Args:
        ttpid (string): the selected ttpid by user from the dropdown list
    """  
    try:
        close_open_file(weekly_dashboard_path)  # Close dashboard if it's open
        logger.info(f"Closed the dashboard file: {weekly_dashboard_path}")

        create_backup(weekly_dashboard_path)    # Create backup of the current file
        logger.info(f"Created backup for file: {weekly_dashboard_path}")

        current_df = create_dataframe_from_sheet(ttpid)  # Create df for the sheet we will be working on
        logger.info(f"Created dataframe for sheet: {ttpid}. Data shape: {current_df.shape}")

        open_file(weekly_dashboard_path)  # Open file after handling
        logger.info(f"Opened the dashboard file: {weekly_dashboard_path}")
    
    except Exception as e:
        logger.error(f"Error during weekly tasks for TTPID {ttpid}: {e}")
        messagebox.showerror(title="Error", message=f"An error occurred during weekly tasks for TTPID {ttpid}: {e}")

def create_dataframe_from_sheet(sheet_name):
    """Create a dataframe from the specified sheet starting from cell A3.
    
    Args:
        sheet_name (string): The name of the sheet to process.
        
    Returns:
        pd.DataFrame: The resulting dataframe with proper headers.
    """
    try:
        logger.info(f"Reading sheet: {sheet_name}")
        
        # Load the sheet into a dataframe, skipping the first two rows
        df = pd.read_excel(weekly_dashboard_path, sheet_name=sheet_name, skiprows=1)
        logger.info(f"Sheet {sheet_name} read successfully. Data shape: {df.shape}")

        # Set the first column as headers
        df.columns = df.iloc[0]
        df = df[1:]  # Remove the header row from the data
        logger.info(f"Headers set from the first column. Data shape after setting headers: {df.shape}")

        return df.reset_index(drop=True)
    except Exception as e:
        logger.error(f"Error processing sheet {sheet_name}: {e}")
        raise e

# ---------------------------- UI SETUP ------------------------------- #
# Main window UI setup
main_window = Tk()
main_window.title("TikTok Weekly Processor")
main_window.config(padx=50, pady=50, bg=WHITE)

# Variable to store the selected TTPID
selected_ttpid = StringVar()
selected_ttpid.set(list(ACTIVE_PRODUCT_LIST.keys())[0])  # Set the default value
# OptionMenu of all available TTPIDs that can be processed
ttpid_optionmenu = OptionMenu(main_window, selected_ttpid, *ACTIVE_PRODUCT_LIST.keys())
ttpid_optionmenu.grid(column=0, row=0)

# Action button that will run the weekly tasks
weekly_task_button = Button(text="Perform Weekly Tasks", command=lambda: weekly_tasks(selected_ttpid.get()))
weekly_task_button.grid(column=1, row=0)

main_window.mainloop()