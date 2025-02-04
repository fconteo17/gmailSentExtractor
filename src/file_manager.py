"""
File management operations for the Gmail Export Tool.
"""
import os
import logging
from datetime import datetime
from src import config

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class FileManager:
    """Handles all file and directory operations for the application."""
    
    def __init__(self):
        """Initialize the FileManager with required directories."""
        self.logger = logging.getLogger(__name__)
        self.create_directory_structure()

    def create_directory_structure(self):
        """Create necessary directories if they don't exist."""
        directories = [
            config.DATA_DIR,
            config.TOKENS_DIR,
            config.EXPORTS_DIR,
            config.CONFIG_DIR,
        ]
        
        for directory in directories:
            try:
                os.makedirs(directory, exist_ok=True)
                self.logger.debug(f"Directory ensured: {directory}")
            except PermissionError:
                self.logger.error(f"Permission denied creating directory: {directory}")
                raise
            except Exception as e:
                self.logger.error(f"Error creating directory {directory}: {str(e)}")
                raise

    def get_token_path(self, email: str) -> str:
        """
        Get path for token file.
        
        Args:
            email: Email address to generate token path for.
            
        Returns:
            str: Full path to the token file.
        """
        if not email:
            raise ValueError("Email address cannot be empty")
            
        filename = f'token_{email.replace("@", "_at_")}.pickle'
        return os.path.join(config.TOKENS_DIR, filename)

    def get_export_path(self, email: str, start_date: datetime) -> str:
        """
        Get path for export file.
        
        Args:
            email: Email address associated with the export.
            start_date: Start date of the export period.
            
        Returns:
            str: Full path to the export file.
        """
        if not email:
            raise ValueError("Email address cannot be empty")
            
        filename = (
            f'sent_emails_{email.replace("@", "_at_")}_'
            f'{start_date.strftime("%Y%m%d")}{config.EXCEL_EXTENSION}'
        )
        return os.path.join(config.EXPORTS_DIR, filename)

    def ensure_credentials_exist(self) -> bool:
        """
        Check if the credentials file exists.
        
        Returns:
            bool: True if credentials exist, False otherwise.
        """
        exists = os.path.exists(config.CREDENTIALS_FILE)
        if not exists:
            self.logger.warning("Credentials file not found!")
        return exists

    def cleanup_token(self, email: str) -> bool:
        """
        Remove token file for a given email.
        
        Args:
            email: Email address whose token should be removed.
            
        Returns:
            bool: True if successful, False otherwise.
        """
        token_path = self.get_token_path(email)
        try:
            if os.path.exists(token_path):
                os.remove(token_path)
                self.logger.info(f"Token removed for {email}")
                return True
        except Exception as e:
            self.logger.error(f"Error removing token for {email}: {str(e)}")
        return False 