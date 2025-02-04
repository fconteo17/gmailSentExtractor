"""
Configuration settings for the Gmail Export Tool.
"""
import os

# Gmail API configuration
GMAIL_API_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

# Directory structure
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
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