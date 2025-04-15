"""
Account management operations for the Gmail Export Tool.
"""
import os
import json
import logging
import re
from typing import List, Dict, Optional
from src import config

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class GmailAccountManager:
    """Manages Gmail account operations."""

    def __init__(self, file_manager):
        """Initialize the account manager."""
        self.logger = logging.getLogger(__name__)
        self.file_manager = file_manager
        self._ensure_accounts_file()

    def _ensure_accounts_file(self):
        """Ensure accounts file exists."""
        if not os.path.exists(config.ACCOUNTS_FILE):
            with open(config.ACCOUNTS_FILE, "w") as f:
                json.dump({"accounts": []}, f)

    def _load_accounts(self) -> Dict:
        """Load accounts from file."""
        try:
            with open(config.ACCOUNTS_FILE, "r") as f:
                return json.load(f)
        except Exception as e:
            self.logger.error(f"Error loading accounts: {str(e)}")
            return {"accounts": []}

    def _save_accounts(self, accounts_data: Dict):
        """Save accounts to file."""
        try:
            with open(config.ACCOUNTS_FILE, "w") as f:
                json.dump(accounts_data, f, indent=4)
        except Exception as e:
            self.logger.error(f"Error saving accounts: {str(e)}")

    def list_accounts(self) -> List[str]:
        """List all registered accounts."""
        accounts_data = self._load_accounts()
        return [acc["email"] for acc in accounts_data.get("accounts", [])]

    def get_account_details(self, email: str) -> Optional[Dict]:
        """Get account details including auth method and app password if IMAP."""
        accounts_data = self._load_accounts()
        for account in accounts_data.get("accounts", []):
            if account["email"] == email:
                return account
        return None

    def add_account(self, email: str, auth_method: str = "oauth", app_password: str = None) -> bool:
        """
        Add a new Gmail account.
        
        Args:
            email: Email address to add.
            auth_method: Authentication method ('oauth' or 'imap').
            app_password: App password for IMAP authentication.
            
        Returns:
            bool: True if successful, False otherwise.
        """
        try:
            accounts_data = self._load_accounts()
            
            # Check if account already exists
            if any(acc["email"] == email for acc in accounts_data["accounts"]):
                self.logger.warning(f"Account {email} already exists")
                return False
            
            # Add new account with auth method
            account_data = {
                "email": email,
                "auth_method": auth_method
            }
            
            # Add app password for IMAP accounts
            if auth_method == "imap":
                if not app_password:
                    raise ValueError("App password required for IMAP authentication")
                account_data["app_password"] = app_password
            
            accounts_data["accounts"].append(account_data)
            self._save_accounts(accounts_data)
            
            self.logger.info(f"Added account {email} with {auth_method} authentication")
            return True
            
        except Exception as e:
            self.logger.error(f"Error adding account {email}: {str(e)}")
            return False

    def remove_account(self, email: str) -> bool:
        """Remove a Gmail account."""
        try:
            accounts_data = self._load_accounts()
            
            # Find and remove account
            accounts_data["accounts"] = [
                acc for acc in accounts_data["accounts"]
                if acc["email"] != email
            ]
            
            self._save_accounts(accounts_data)
            
            # Clean up token file if exists
            self.file_manager.cleanup_token(email)
            
            self.logger.info(f"Removed account {email}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error removing account {email}: {str(e)}")
            return False

    def get_account_token_path(self, email: str) -> Optional[str]:
        """Get token path for OAuth accounts only."""
        account = self.get_account_details(email)
        if account and account.get("auth_method") == "oauth":
            return self.file_manager.get_token_path(email)
        return None

    @staticmethod
    def _validate_email(email: str) -> bool:
        """
        Validate email format.
        
        Args:
            email: Email address to validate.
            
        Returns:
            bool: True if valid, False otherwise.
        """
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email)) 