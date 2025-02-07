"""
Configuration settings for the Gmail Export Tool.
"""
import os
import sys

# Gmail API configuration
GMAIL_API_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

def get_base_path():
    """Get the base path for the application in both dev and PyInstaller environments."""
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle (PyInstaller)
        base_path = os.path.dirname(sys.executable)
    else:
        # If the application is run from a Python interpreter
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return base_path

# Directory structure
BASE_DIR = get_base_path()
DATA_DIR = os.path.join(BASE_DIR, "data")
TOKENS_DIR = os.path.join(DATA_DIR, "tokens")
EXPORTS_DIR = os.path.join(DATA_DIR, "exports")
CONFIG_DIR = os.path.join(DATA_DIR, "config")

# File paths
ACCOUNTS_FILE = os.path.join(CONFIG_DIR, "email_accounts.json")
CREDENTIALS_FILE = os.path.join(CONFIG_DIR, "credentials.json")

# Date format for user input
DATE_FORMAT = "%Y-%m-%d"
DATE_FORMAT_DISPLAY = "YYYY-MM-DD"

# Export settings
EXCEL_EXTENSION = ".xlsx"
EXCEL_FILENAME_FORMAT = "sent_emails_{start_date}_{end_date}"  # Will be formatted with dates 