"""
Gmail API service operations for the Gmail Export Tool.
"""
import os
import pickle
import logging
import re
from datetime import datetime
from typing import List, Dict, Any, Tuple
from tqdm import tqdm

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from src import config

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class GmailService:
    """Handles all Gmail API operations."""

    def __init__(self, file_manager, account_manager):
        """
        Initialize the Gmail service.
        
        Args:
            file_manager: Instance of FileManager for file operations.
            account_manager: Instance of GmailAccountManager for account operations.
        """
        self.logger = logging.getLogger(__name__)
        self.file_manager = file_manager
        self.account_manager = account_manager

    @staticmethod
    def _split_email(email: str) -> Tuple[str, str]:
        """
        Split email address into username and domain.
        
        Args:
            email: Email address to split.
            
        Returns:
            Tuple[str, str]: Username and domain parts.
        """
        try:
            # Handle multiple email addresses
            if ',' in email:
                # Take only the first email address
                email = email.split(',')[0].strip()
            
            # Extract email from format like '"Name" <email@domain.com>'
            email_pattern = r'<?([^<]*@[^>]*)>?'
            match = re.search(email_pattern, email)
            if match:
                email = match.group(1).strip()
            
            # Split at @ symbol
            parts = email.split('@')
            if len(parts) == 2:
                username = parts[0].strip()
                domain = parts[1].strip()
                return username, domain
            return email, ""  # Return original as username if not valid email format
        except Exception:
            return email, ""  # Return original as username in case of any error

    def setup_service(self, email: str):
        """
        Set up Gmail API service for a specific email.
        
        Args:
            email: Email address to set up service for.
            
        Returns:
            Resource: Gmail API service resource.
            
        Raises:
            FileNotFoundError: If credentials file is missing.
            ValueError: If email is invalid or not registered.
        """
        if not self.account_manager.get_account_token_path(email):
            raise ValueError(f"Email {email} not registered")

        creds = None
        token_file = self.account_manager.get_account_token_path(email)

        # Load existing credentials
        if os.path.exists(token_file):
            try:
                with open(token_file, "rb") as token:
                    creds = pickle.load(token)
            except Exception as e:
                self.logger.error(f"Error loading credentials: {str(e)}")

        # Refresh or create new credentials
        try:
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    if not self.file_manager.ensure_credentials_exist():
                        raise FileNotFoundError(
                            "credentials.json not found in config directory"
                        )
                    
                    flow = InstalledAppFlow.from_client_secrets_file(
                        config.CREDENTIALS_FILE,
                        config.GMAIL_API_SCOPES
                    )
                    creds = flow.run_local_server(port=0)

                # Save credentials
                with open(token_file, "wb") as token:
                    pickle.dump(creds, token)
        except Exception as e:
            self.logger.error(f"Error setting up credentials: {str(e)}")
            raise

        return build("gmail", "v1", credentials=creds)

    def get_sent_emails(
        self,
        service: Any,
        start_date: datetime,
        end_date: datetime
    ) -> List[Dict]:
        """
        Fetch sent emails between specified dates.
        
        Args:
            service: Gmail API service instance.
            start_date: Start date for email fetch.
            end_date: End date for email fetch.
            
        Returns:
            List[Dict]: List of email data dictionaries.
        """
        # Format dates for Gmail query
        start_str = start_date.strftime("%Y/%m/%d")
        end_str = end_date.strftime("%Y/%m/%d")
        query = f"in:sent after:{start_str} before:{end_str}"

        try:
            # Get list of messages
            results = service.users().messages().list(userId="me", q=query).execute()
            messages = results.get("messages", [])

            if not messages:
                self.logger.info("No emails found in the specified date range")
                return []

            emails_data = []
            with tqdm(total=len(messages), desc="Fetching emails") as pbar:
                for message in messages:
                    try:
                        msg = (
                            service.users()
                            .messages()
                            .get(userId="me", id=message["id"])
                            .execute()
                        )
                        
                        headers = msg["payload"]["headers"]
                        to_address = self._get_header(headers, "to", "No Recipient")
                        username, domain = self._split_email(to_address)
                        
                        email_data = {
                            "Date": self._get_header(headers, "date", "No Date"),
                            "Username": username,
                            "Domain": domain,
                            "Subject": self._get_header(headers, "subject", "No Subject")
                        }
                        emails_data.append(email_data)
                        pbar.update(1)
                    except HttpError as e:
                        self.logger.error(f"Error fetching message {message['id']}: {str(e)}")
                    except Exception as e:
                        self.logger.error(f"Unexpected error processing message: {str(e)}")

            return emails_data

        except HttpError as e:
            self.logger.error(f"Error accessing Gmail API: {str(e)}")
            raise
        except Exception as e:
            self.logger.error(f"Unexpected error: {str(e)}")
            raise

    @staticmethod
    def _get_header(headers: List[Dict], name: str, default: str = "") -> str:
        """
        Get header value from email headers.
        
        Args:
            headers: List of email headers.
            name: Name of the header to find.
            default: Default value if header not found.
            
        Returns:
            str: Header value or default.
        """
        return next(
            (h["value"] for h in headers if h["name"].lower() == name.lower()),
            default
        ) 