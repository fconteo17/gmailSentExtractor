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
    """Manages Gmail account configurations and credentials."""

    def __init__(self, file_manager):
        """
        Initialize the account manager.
        
        Args:
            file_manager: Instance of FileManager for file operations.
        """
        self.logger = logging.getLogger(__name__)
        self.file_manager = file_manager
        self.accounts = self._load_accounts()

    def _load_accounts(self) -> Dict:
        """
        Load saved email accounts from configuration file.
        
        Returns:
            dict: Dictionary of saved accounts.
        """
        try:
            if os.path.exists(config.ACCOUNTS_FILE):
                with open(config.ACCOUNTS_FILE, "r") as f:
                    return json.load(f)
        except json.JSONDecodeError:
            self.logger.error("Invalid accounts file format")
        except Exception as e:
            self.logger.error(f"Error loading accounts: {str(e)}")
        return {}

    def _save_accounts(self) -> bool:
        """
        Save accounts to configuration file.
        
        Returns:
            bool: True if successful, False otherwise.
        """
        try:
            with open(config.ACCOUNTS_FILE, "w") as f:
                json.dump(self.accounts, f, indent=4)
            return True
        except Exception as e:
            self.logger.error(f"Error saving accounts: {str(e)}")
            return False

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

    def add_account(self, email: str) -> bool:
        """
        Add a new email account.
        
        Args:
            email: Email address to add.
            
        Returns:
            bool: True if successful, False otherwise.
        """
        if not self._validate_email(email):
            self.logger.error(f"Invalid email format: {email}")
            return False

        if email in self.accounts:
            self.logger.warning(f"Account {email} already exists")
            return False

        try:
            token_file = self.file_manager.get_token_path(email)
            self.accounts[email] = {"token_file": token_file}
            return self._save_accounts()
        except Exception as e:
            self.logger.error(f"Error adding account {email}: {str(e)}")
            return False

    def remove_account(self, email: str) -> bool:
        """
        Remove an email account.
        
        Args:
            email: Email address to remove.
            
        Returns:
            bool: True if successful, False otherwise.
        """
        if email not in self.accounts:
            self.logger.warning(f"Account {email} not found")
            return False

        try:
            # Remove token file
            self.file_manager.cleanup_token(email)
            
            # Remove from accounts
            del self.accounts[email]
            return self._save_accounts()
        except Exception as e:
            self.logger.error(f"Error removing account {email}: {str(e)}")
            return False

    def get_account_token_path(self, email: str) -> Optional[str]:
        """
        Get token file path for an account.
        
        Args:
            email: Email address to get token for.
            
        Returns:
            Optional[str]: Token file path if account exists, None otherwise.
        """
        account = self.accounts.get(email)
        return account["token_file"] if account else None

    def list_accounts(self) -> List[str]:
        """
        Get list of all saved accounts.
        
        Returns:
            List[str]: List of email addresses.
        """
        return list(self.accounts.keys()) 