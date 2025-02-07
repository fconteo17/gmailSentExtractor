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
from src.cli import EnhancedCLI

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
        self.ui = EnhancedCLI()

    def export_single_account(self, email: str, start_date: datetime, end_date: datetime) -> bool:
        """Export emails from a single account."""
        try:
            # Setup service
            self.ui.start_operation(f"Authenticating {email}")
            service = self.gmail_service.setup_service(email)
            self.ui.stop_operation()

            # Get emails
            self.ui.display_success(f"Starting email fetch for {email}...")
            emails_data = []
            
            # Get total count first
            total_messages = self.gmail_service.get_total_messages(
                service, start_date, end_date
            )
            
            if total_messages == 0:
                self.ui.display_error(f"No emails found for {email} in the specified date range")
                return False

            # Create progress bar
            with self.ui.display_progress(
                total=total_messages,
                desc=f"Fetching emails from {email}"
            ) as pbar:
                for email_batch in self.gmail_service.get_sent_emails_with_progress(
                    service, start_date, end_date
                ):
                    emails_data.extend(email_batch)
                    pbar.update(1)

            if not emails_data:
                self.ui.display_error(f"No emails found for {email} in the specified date range")
                return False

            # Export to Excel
            self.ui.start_operation(f"Exporting {email} to Excel...")
            output_file = self.file_manager.get_export_path(email, start_date, end_date)
            if self.export_manager.export_to_excel(emails_data, output_file, email):
                self.ui.display_success(f"Successfully exported {len(emails_data)} emails from {email} to {output_file}")
                return True
            else:
                self.ui.display_error(f"Failed to export emails for {email}")
                return False

        except Exception as e:
            self.ui.display_error(f"Error processing {email}: {str(e)}")
            return False
        finally:
            self.ui.stop_operation()

    def run(self):
        """Run the main application loop."""
        while True:
            self.ui.display_banner()
            self.ui.display_menu()
            choice = self.ui.get_menu_choice()

            if choice == 1:  # Add Gmail Account
                email = self.ui.get_email()
                if email:
                    try:
                        self.ui.start_operation(f"Setting up {email}")
                        service = self.gmail_service.setup_service(email)
                        if service:
                            self.account_manager.add_account(email)
                            self.ui.display_success(f"Successfully added {email}")
                    except Exception as e:
                        self.ui.display_error(str(e))
                    finally:
                        self.ui.stop_operation()

            elif choice == 2:  # Remove Gmail Account
                accounts = self.account_manager.list_accounts()
                email = self.ui.select_account(accounts)
                if email:
                    if self.ui.confirm_action(f"Remove {email}?"):
                        if self.account_manager.remove_account(email):
                            self.ui.display_success(f"Successfully removed {email}")
                        else:
                            self.ui.display_error(f"Failed to remove {email}")

            elif choice == 3:  # Export Emails (Single Account)
                accounts = self.account_manager.list_accounts()
                if not accounts:
                    self.ui.display_error("No accounts available. Please add an account first.")
                    continue

                email = self.ui.select_account(accounts)
                if not email:
                    continue

                try:
                    start_date, end_date = self.ui.get_date_range()
                    self.export_single_account(email, start_date, end_date)
                except Exception as e:
                    self.ui.display_error(str(e))

            elif choice == 4:  # Export Emails (All Accounts)
                accounts = self.account_manager.list_accounts()
                if not accounts:
                    self.ui.display_error("No accounts available. Please add an account first.")
                    continue

                if not self.ui.confirm_action("Export emails from all accounts?"):
                    continue

                try:
                    start_date, end_date = self.ui.get_date_range()
                    successful_exports = 0
                    failed_exports = 0

                    for email in accounts:
                        if self.export_single_account(email, start_date, end_date):
                            successful_exports += 1
                        else:
                            failed_exports += 1

                    # Display summary
                    if successful_exports > 0:
                        self.ui.display_success(f"Successfully exported emails from {successful_exports} account(s)")
                    if failed_exports > 0:
                        self.ui.display_error(f"Failed to export emails from {failed_exports} account(s)")

                except Exception as e:
                    self.ui.display_error(str(e))

            elif choice == 5:  # Configure Date Range
                try:
                    self.ui.configure_date_range()
                except Exception as e:
                    self.ui.display_error(f"Error configuring dates: {str(e)}")

            elif choice == 6:  # Exit
                if self.ui.confirm_action("Are you sure you want to exit?"):
                    self.ui.display_success("Goodbye!")
                    break


def main():
    """Main entry point."""
    tool = GmailExportTool()
    tool.run()


if __name__ == "__main__":
    main()
