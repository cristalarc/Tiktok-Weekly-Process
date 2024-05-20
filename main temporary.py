import os
import logging
import time
from tkinter import *
from tkinter import messagebox
import subprocess # For file opening
import win32com.client # For detecting if file is open and then closing it
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
# ---------------------------- CONSTANTS ------------------------------- #
FONT_NAME = "Calibri"
FONT_SIZE = 11
WHITE = "#fcf7f9"

# ---------------------------- INPUT CHECKERS ------------------------------- #

# ---------------------------- WEEKLY TASKS -------------------------------- #
def weekly_tasks():
    """Runs the weekly tasks as defined in the instructions.
    These are activated via a button in the UI.\n

    """    
    pass
# ---------------------------- UI SETUP ------------------------------- #
# Main window UI setup
main_window = Tk()
main_window.title("TikTok Weekly Processor")
main_window.config(padx=50, pady=50, bg=WHITE)

# Action button that will run the weekly tasks
weekly_task_button = Button(text="Perform Weekly Tasks", command=weekly_tasks)
weekly_task_button.grid(column=0, row=0)

main_window.mainloop()