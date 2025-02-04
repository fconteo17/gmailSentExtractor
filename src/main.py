"""
Main entry point for the Gmail Export Tool.
"""

import logging
from typing import Optional, Tuple
from datetime import datetime

from src.file_manager import FileManager
from src.account_manager import GmailAccountManager
from src.gmail_service import GmailService
from src.export import ExportManager
from src.ui import UserInterface

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


class GmailExportTool:
    """Main application class coordinating all operations."""

    def __init__(self):
        """Initialize the application components."""
        self.logger = logging.getLogger(__name__)
        self.file_manager = FileManager()
        self.account_manager = GmailAccountManager(self.file_manager)
        self.gmail_service = GmailService(self.file_manager, self.account_manager)
        self.export_manager = ExportManager()
        self.ui = UserInterface()

    def add_account(self) -> bool:
        """
        Handle adding a new email account and automatically fetch emails.

        Returns:
            bool: True if successful, False otherwise.
        """
        email = self.ui.get_email_input()

        if self.account_manager.add_account(email):
            try:
                # Test authentication
                service = self.gmail_service.setup_service(email)
                self.ui.display_success(
                    f"Account {email} added and authenticated successfully"
                )

                # Ask if user wants to fetch emails immediately
                if self.ui.confirm_action(
                    "Would you like to fetch emails for this account now?"
                ):
                    try:
                        # Get date range
                        start_date, end_date = self.ui.get_date_range()

                        # Get emails
                        self.ui.display_success("Fetching emails...")
                        emails_data = self.gmail_service.get_sent_emails(
                            service, start_date, end_date
                        )

                        if not emails_data:
                            self.ui.display_error(
                                "No emails found in the specified date range"
                            )
                            return True  # Account was still added successfully

                        # Export to Excel
                        output_file = self.file_manager.get_export_path(
                            email, start_date
                        )
                        if self.export_manager.export_to_excel(
                            emails_data, output_file, email
                        ):
                            self.ui.display_export_summary(
                                email, len(emails_data), output_file
                            )
                            return True

                        self.ui.display_error("Failed to export emails")
                        return True  # Account was still added successfully

                    except Exception as e:
                        self.ui.display_error(f"Error fetching emails: {str(e)}")
                        return True  # Account was still added successfully

                return True

            except Exception as e:
                self.account_manager.remove_account(email)
                self.ui.display_error(f"Authentication failed: {str(e)}")
                return False
        return False

    def use_existing_account(self) -> bool:
        """
        Handle using an existing account for export.

        Returns:
            bool: True if successful, False otherwise.
        """
        accounts = self.account_manager.list_accounts()
        email = self.ui.select_account(accounts)

        if not email:
            return False

        try:
            # Get date range
            start_date, end_date = self.ui.get_date_range()

            # Setup service
            service = self.gmail_service.setup_service(email)

            # Get emails
            self.ui.display_success("Fetching emails...")
            emails_data = self.gmail_service.get_sent_emails(
                service, start_date, end_date
            )

            if not emails_data:
                self.ui.display_error("No emails found in the specified date range")
                return False

            # Export to Excel
            output_file = self.file_manager.get_export_path(email, start_date)
            if self.export_manager.export_to_excel(emails_data, output_file, email):
                self.ui.display_export_summary(email, len(emails_data), output_file)
                return True

            self.ui.display_error("Failed to export emails")
            return False

        except Exception as e:
            self.ui.display_error(str(e))
            return False

    def remove_account(self) -> bool:
        """
        Handle removing an email account.

        Returns:
            bool: True if successful, False otherwise.
        """
        accounts = self.account_manager.list_accounts()
        email = self.ui.select_account(accounts)

        if not email:
            return False

        if self.ui.confirm_action(f"Are you sure you want to remove {email}?"):
            if self.account_manager.remove_account(email):
                self.ui.display_success(f"Account {email} removed successfully")
                return True
            self.ui.display_error(f"Failed to remove account {email}")
        return False

    def run(self):
        """Run the main application loop."""
        while True:
            try:
                choice = self.ui.display_menu()

                if choice == "1":
                    self.add_account()
                elif choice == "2":
                    self.use_existing_account()
                elif choice == "3":
                    self.remove_account()
                elif choice == "4":
                    self.ui.display_success("Goodbye!")
                    break

            except KeyboardInterrupt:
                self.ui.display_success("\nGoodbye!")
                break
            except Exception as e:
                self.ui.display_error(f"Unexpected error: {str(e)}")


def main():
    """Application entry point."""
    try:
        app = GmailExportTool()
        app.run()
    except Exception as e:
        logging.error(f"Application error: {str(e)}")
        raise


if __name__ == "__main__":
    main()
