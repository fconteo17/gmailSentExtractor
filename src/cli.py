"""
Enhanced CLI interface for the Gmail Export Tool.
"""
import sys
from datetime import datetime
from typing import Optional, Tuple, List, Union, Iterator
import re
from colorama import init, Fore, Style, Cursor
from halo import Halo
from tqdm import tqdm

# Initialize colorama for Windows support
init()

class EnhancedCLI:
    """Enhanced CLI interface with colors and interactive features."""

    def __init__(self):
        """Initialize the CLI interface."""
        self.spinner = Halo(spinner='dots')
        self.current_progress_bar = None

    def display_banner(self):
        """Display a colorful banner."""
        banner = f"""
{Fore.CYAN}╔══════════════════════════════════════╗
║     Gmail Sent Emails Export Tool      ║
╚══════════════════════════════════════╝{Style.RESET_ALL}
"""
        print(banner)

    def display_menu(self) -> str:
        """
        Display the main menu with colors and get user choice.
        
        Returns:
            str: User's menu choice.
        """
        self.display_banner()
        print(f"\n{Fore.YELLOW}Available Options:{Style.RESET_ALL}")
        print(f"{Fore.GREEN}1.{Style.RESET_ALL} Add new email account")
        print(f"{Fore.GREEN}2.{Style.RESET_ALL} Use existing account")
        print(f"{Fore.GREEN}3.{Style.RESET_ALL} Remove email account")
        print(f"{Fore.GREEN}4.{Style.RESET_ALL} Exit")
        
        while True:
            choice = input(f"\n{Fore.CYAN}Select an option (1-4):{Style.RESET_ALL} ").strip()
            if choice in ["1", "2", "3", "4"]:
                return choice
            print(f"{Fore.RED}Invalid choice. Please select 1-4.{Style.RESET_ALL}")

    def get_email_input(self) -> str:
        """
        Get email input from user with validation.
        
        Returns:
            str: User's email input.
        """
        while True:
            email = input(f"\n{Fore.CYAN}Enter email address:{Style.RESET_ALL} ").strip()
            if self._validate_email(email):
                return email
            print(f"{Fore.RED}Invalid email format. Please try again.{Style.RESET_ALL}")

    def _validate_email(self, email: str) -> bool:
        """Validate email format."""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email))

    def select_account(self, accounts: list) -> Optional[str]:
        """
        Let user select an account from the list with colored output.
        
        Args:
            accounts: List of available email accounts.
            
        Returns:
            Optional[str]: Selected email or None if no selection made.
        """
        if not accounts:
            print(f"{Fore.RED}No email accounts saved.{Style.RESET_ALL}")
            return None

        print(f"\n{Fore.YELLOW}Saved email accounts:{Style.RESET_ALL}")
        for i, email in enumerate(accounts, 1):
            print(f"{Fore.GREEN}{i}.{Style.RESET_ALL} {email}")

        while True:
            try:
                idx = input(f"\n{Fore.CYAN}Select account number (or 0 to cancel):{Style.RESET_ALL} ")
                idx = int(idx)
                if idx == 0:
                    return None
                if 1 <= idx <= len(accounts):
                    return accounts[idx - 1]
                print(f"{Fore.RED}Please enter a number between 1 and {len(accounts)}{Style.RESET_ALL}")
            except ValueError:
                print(f"{Fore.RED}Please enter a valid number{Style.RESET_ALL}")

    def get_date_range(self) -> Tuple[datetime, datetime]:
        """
        Get date range input from user with improved validation.
        
        Returns:
            Tuple[datetime, datetime]: Start and end dates.
        """
        date_format = "%Y-%m-%d"
        format_display = "YYYY-MM-DD"
        
        print(f"\n{Fore.YELLOW}Enter dates in format: {format_display}{Style.RESET_ALL}")
        
        while True:
            try:
                start_str = input(f"{Fore.CYAN}Start date:{Style.RESET_ALL} ").strip()
                start_date = datetime.strptime(start_str, date_format)
                break
            except ValueError:
                print(f"{Fore.RED}Invalid date format. Please use {format_display}{Style.RESET_ALL}")

        while True:
            try:
                end_str = input(f"{Fore.CYAN}End date:{Style.RESET_ALL} ").strip()
                end_date = datetime.strptime(end_str, date_format)
                
                if end_date < start_date:
                    print(f"{Fore.RED}End date must be after start date{Style.RESET_ALL}")
                    continue
                    
                break
            except ValueError:
                print(f"{Fore.RED}Invalid date format. Please use {format_display}{Style.RESET_ALL}")

        return start_date, end_date

    def confirm_action(self, message: str) -> bool:
        """
        Get user confirmation for an action with colored prompt.
        
        Args:
            message: Confirmation message to display.
            
        Returns:
            bool: True if confirmed, False otherwise.
        """
        response = input(f"\n{Fore.YELLOW}{message} (y/n):{Style.RESET_ALL} ").strip().lower()
        return response == "y"

    def display_progress(
        self,
        iterable: Optional[Iterator] = None,
        desc: Optional[str] = None,
        total: Optional[int] = None
    ) -> Union[tqdm, Iterator]:
        """
        Display progress information using tqdm.
        
        Args:
            iterable: Optional iterable to track progress of.
            desc: Optional description of the progress.
            total: Optional total for manual progress tracking.
            
        Returns:
            Union[tqdm, Iterator]: Progress bar or wrapped iterator.
        """
        # Stop any running spinner to prevent conflicts
        self.stop_operation()
        
        # Create and store the progress bar with a simpler format
        self.current_progress_bar = tqdm(
            iterable=iterable,
            desc=desc,
            total=total,
            unit="emails",
            leave=True,
            dynamic_ncols=True,
            miniters=1
        )
        
        return self.current_progress_bar

    def start_operation(self, message: str):
        """
        Start a spinner for an operation.
        
        Args:
            message: Message to display during the operation.
        """
        # Stop any existing progress bar to prevent conflicts
        if self.current_progress_bar is not None:
            self.current_progress_bar.close()
            self.current_progress_bar = None
            
        self.spinner.text = message
        self.spinner.start()

    def stop_operation(self):
        """Stop the current spinner."""
        if self.spinner.spinner_id is not None:
            self.spinner.stop()
        
        # Also clear any existing progress bar
        if self.current_progress_bar is not None:
            self.current_progress_bar.close()
            self.current_progress_bar = None

    def display_error(self, message: str):
        """
        Display error message to user with red color.
        
        Args:
            message: Error message to display.
        """
        print(f"\n{Fore.RED}Error: {message}{Style.RESET_ALL}")

    def display_success(self, message: str):
        """
        Display success message to user with green color.
        
        Args:
            message: Success message to display.
        """
        print(f"\n{Fore.GREEN}Success: {message}{Style.RESET_ALL}")

    def display_export_summary(self, email: str, count: int, file_path: str):
        """
        Display export operation summary with colors.
        
        Args:
            email: Email address used for export.
            count: Number of emails exported.
            file_path: Path to export file.
        """
        print(f"\n{Fore.YELLOW}Export Summary:{Style.RESET_ALL}")
        print(f"{Fore.CYAN}Email Account:{Style.RESET_ALL} {email}")
        print(f"{Fore.CYAN}Emails Exported:{Style.RESET_ALL} {count}")
        print(f"{Fore.CYAN}Export File:{Style.RESET_ALL} {file_path}") 