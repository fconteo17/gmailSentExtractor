"""
Gmail API service operations for the Gmail Export Tool.
"""

import os
import pickle
import logging
import re
from datetime import datetime
from typing import List, Dict, Any, Tuple, Generator
from tqdm import tqdm
import base64
from email.utils import parsedate_to_datetime

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from src import config

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
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
            if "," in email:
                # Take only the first email address
                email = email.split(",")[0].strip()

            # Extract email from format like '"Name" <email@domain.com>'
            email_pattern = r"<?([^<]*@[^>]*)>?"
            match = re.search(email_pattern, email)
            if match:
                email = match.group(1).strip()

            # Split at @ symbol
            parts = email.split("@")
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
                        config.CREDENTIALS_FILE, config.GMAIL_API_SCOPES
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
        self, service: Any, start_date: datetime, end_date: datetime
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
                            "Subject": self._get_header(
                                headers, "subject", "No Subject"
                            ),
                        }
                        emails_data.append(email_data)
                        pbar.update(1)
                    except HttpError as e:
                        self.logger.error(
                            f"Error fetching message {message['id']}: {str(e)}"
                        )
                    except Exception as e:
                        self.logger.error(
                            f"Unexpected error processing message: {str(e)}"
                        )

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
            (h["value"] for h in headers if h["name"].lower() == name.lower()), default
        )

    def get_total_messages(
        self, service: Any, start_date: datetime, end_date: datetime
    ) -> int:
        """
        Get total count of messages in the specified date range.

        Args:
            service: Gmail API service instance.
            start_date: Start date for email range.
            end_date: End date for email range.

        Returns:
            int: Total number of messages.
        """
        query = f'in:sent after:{start_date.strftime("%Y/%m/%d")} before:{end_date.strftime("%Y/%m/%d")}'

        try:
            # Get all message IDs first to get accurate count
            result = service.users().messages().list(userId="me", q=query).execute()

            messages = result.get("messages", [])
            total = len(messages)

            # Get remaining messages if there are more pages
            while "nextPageToken" in result:
                result = (
                    service.users()
                    .messages()
                    .list(userId="me", q=query, pageToken=result["nextPageToken"])
                    .execute()
                )
                messages = result.get("messages", [])
                total += len(messages)

            return total

        except Exception as e:
            self.logger.error(f"Error getting message count: {str(e)}")
            return 0

    def get_sent_emails_with_progress(
        self, service, start_date: datetime, end_date: datetime
    ) -> Generator[List[Dict[str, Any]], None, None]:
        """
        Get sent emails with progress reporting.

        Args:
            service: Gmail API service instance.
            start_date: Start date for email range.
            end_date: End date for email range.

        Yields:
            List[Dict[str, Any]]: Batches of email data.
        """
        query = f'in:sent after:{start_date.strftime("%Y/%m/%d")} before:{end_date.strftime("%Y/%m/%d")}'
        page_token = None
        processed_count = 0

        while True:
            # Get batch of message IDs
            result = (
                service.users()
                .messages()
                .list(
                    userId="me",
                    q=query,
                    pageToken=page_token,
                    maxResults=100,  # Process in batches of 100
                )
                .execute()
            )

            messages = result.get("messages", [])
            if not messages:
                break

            # Process each message in the batch
            batch_data = []
            for message in messages:
                try:
                    msg = (
                        service.users()
                        .messages()
                        .get(userId="me", id=message["id"], format="full")
                        .execute()
                    )

                    # Extract email data
                    headers = msg["payload"]["headers"]
                    to_address = self._get_header(headers, "to", "No Recipient")
                    username, domain = self._split_email(to_address)

                    # Get the date and verify it's within range
                    date_str = self._get_header(headers, "date", None)
                    if not date_str:
                        self.logger.error(f"No date found for message {message['id']}")
                        continue

                    try:
                        # Keep the original date string for pandas to parse
                        email_data = {
                            "Date": date_str,  # Keep the original RFC 2822 date string
                            "Username": username,
                            "Domain": domain,
                            "Subject": self._get_header(
                                headers, "subject", "No Subject"
                            ),
                        }

                        batch_data.append(email_data)
                        processed_count += 1

                        # Yield single-item batches for more accurate progress
                        yield [email_data]

                    except Exception as e:
                        self.logger.error(f"Error parsing date '{date_str}': {str(e)}")
                        continue  # Skip emails with invalid dates

                except Exception as e:
                    self.logger.error(
                        f"Error processing message {message['id']}: {str(e)}"
                    )
                    continue

            # Get the next page token
            page_token = result.get("nextPageToken")
            if not page_token:
                break

    def _get_body_from_parts(self, parts: List[Dict[str, Any]]) -> str:
        """
        Extract email body from message parts.

        Args:
            parts: List of message parts.

        Returns:
            str: Email body text.
        """
        text = []
        for part in parts:
            if part.get("mimeType") == "text/plain":
                text.append(
                    base64.urlsafe_b64decode(
                        part["body"].get("data", "").encode("utf-8")
                    ).decode("utf-8")
                )
            elif "parts" in part:
                text.append(self._get_body_from_parts(part["parts"]))
        return "\n".join(text)
