"""
User interface operations for the Gmail Export Tool.
"""
import logging
from datetime import datetime
from typing import Tuple, Optional
from src import config

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class UserInterface:
    """Handles all user interface operations."""

    def __init__(self):
        """Initialize the UserInterface."""
        self.logger = logging.getLogger(__name__)

    def display_menu(self) -> str:
        """
        Display the main menu and get user choice.
        
        Returns:
            str: User's menu choice.
        """
        print("\n=== Gmail Export Tool ===")
        print("1. Add new email account")
        print("2. Use existing account")
        print("3. Remove email account")
        print("4. Exit")
        
        while True:
            choice = input("Select an option (1-4): ").strip()
            if choice in ["1", "2", "3", "4"]:
                return choice
            print("Invalid choice. Please select 1-4.")

    def get_email_input(self) -> str:
        """
        Get email input from user.
        
        Returns:
            str: User's email input.
        """
        return input("Enter email address: ").strip()

    def select_account(self, accounts: list) -> Optional[str]:
        """
        Let user select an account from the list.
        
        Args:
            accounts: List of available email accounts.
            
        Returns:
            Optional[str]: Selected email or None if no selection made.
        """
        if not accounts:
            print("No email accounts saved.")
            return None

        print("\nSaved email accounts:")
        for i, email in enumerate(accounts, 1):
            print(f"{i}. {email}")

        while True:
            try:
                idx = int(input("Select account number (or 0 to cancel): "))
                if idx == 0:
                    return None
                if 1 <= idx <= len(accounts):
                    return accounts[idx - 1]
                print(f"Please enter a number between 1 and {len(accounts)}")
            except ValueError:
                print("Please enter a valid number")

    def get_date_range(self) -> Tuple[datetime, datetime]:
        """
        Get date range input from user.
        
        Returns:
            Tuple[datetime, datetime]: Start and end dates.
            
        Raises:
            ValueError: If invalid date format or range.
        """
        while True:
            try:
                print(f"\nEnter dates in format: {config.DATE_FORMAT_DISPLAY}")
                start_str = input("Start date: ").strip()
                end_str = input("End date: ").strip()

                start_date = datetime.strptime(start_str, config.DATE_FORMAT)
                end_date = datetime.strptime(end_str, config.DATE_FORMAT)

                if end_date < start_date:
                    print("End date must be after start date")
                    continue

                return start_date, end_date
            except ValueError:
                print(f"Invalid date format. Please use {config.DATE_FORMAT_DISPLAY}")

    def confirm_action(self, message: str) -> bool:
        """
        Get user confirmation for an action.
        
        Args:
            message: Confirmation message to display.
            
        Returns:
            bool: True if confirmed, False otherwise.
        """
        response = input(f"{message} (y/n): ").strip().lower()
        return response == "y"

    def display_progress(self, current: int, total: int, message: str):
        """
        Display progress information.
        
        Args:
            current: Current progress value.
            total: Total progress value.
            message: Progress message to display.
        """
        percentage = (current / total) * 100 if total > 0 else 0
        print(f"\r{message}: {percentage:.1f}% ({current}/{total})", end="")

    def display_error(self, message: str):
        """
        Display error message to user.
        
        Args:
            message: Error message to display.
        """
        print(f"\nError: {message}")

    def display_success(self, message: str):
        """
        Display success message to user.
        
        Args:
            message: Success message to display.
        """
        print(f"\nSuccess: {message}")

    def display_export_summary(self, email: str, count: int, file_path: str):
        """
        Display export operation summary.
        
        Args:
            email: Email address used for export.
            count: Number of emails exported.
            file_path: Path to export file.
        """
        print("\nExport Summary:")
        print(f"Email Account: {email}")
        print(f"Emails Exported: {count}")
        print(f"Export File: {file_path}") 