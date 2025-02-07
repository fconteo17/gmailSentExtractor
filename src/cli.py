"""
Enhanced CLI interface for Gmail Export Tool.
"""
import json
import os
from datetime import datetime
from typing import Tuple, Optional
from colorama import Fore, Style, init
from halo import Halo
from tqdm import tqdm
from src import config

# Initialize colorama
init()

class EnhancedCLI:
    """Enhanced CLI interface with colors and interactive features."""

    def __init__(self):
        """Initialize the CLI interface."""
        self.spinner = None
        self._ensure_config_files()

    def _ensure_config_files(self):
        """Ensure necessary config files exist."""
        date_config_file = os.path.join(config.CONFIG_DIR, "date_config.json")
        if not os.path.exists(date_config_file):
            default_config = {
                "start_date": datetime.now().replace(day=1).strftime("%Y-%m-%d"),
                "end_date": datetime.now().strftime("%Y-%m-%d")
            }
            with open(date_config_file, "w") as f:
                json.dump(default_config, f, indent=4)

    def display_banner(self):
        """Display the application banner."""
        print(f"\n{Fore.CYAN}Gmail Export Tool{Style.RESET_ALL}")
        print("=" * 50)

    def display_menu(self):
        """Display the main menu."""
        print(f"\n{Fore.YELLOW}Menu Options:{Style.RESET_ALL}")
        print("1. Add Gmail Account")
        print("2. Remove Gmail Account")
        print("3. Export Emails (Single Account)")
        print("4. Export Emails (All Accounts)")
        print("5. Configure Date Range")
        print("6. Exit")

    def get_menu_choice(self) -> int:
        """
        Get user's menu choice.
        
        Returns:
            int: Selected menu option.
        """
        while True:
            try:
                choice = int(input(f"\n{Fore.CYAN}Enter your choice (1-6):{Style.RESET_ALL} "))
                if 1 <= choice <= 6:
                    return choice
                print(f"{Fore.RED}Please enter a number between 1 and 6{Style.RESET_ALL}")
            except ValueError:
                print(f"{Fore.RED}Please enter a valid number{Style.RESET_ALL}")

    def get_email(self) -> str:
        """
        Get email input from user.
        
        Returns:
            str: Email address.
        """
        return input(f"\n{Fore.CYAN}Enter Gmail address:{Style.RESET_ALL} ").strip()

    def select_account(self, accounts: list) -> Optional[str]:
        """
        Let user select an account from the list.
        
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
            print(f"{i}. {email}")

        while True:
            try:
                idx = int(input(f"\n{Fore.CYAN}Select account number (or 0 to cancel):{Style.RESET_ALL} "))
                if idx == 0:
                    return None
                if 1 <= idx <= len(accounts):
                    return accounts[idx - 1]
                print(f"{Fore.RED}Please enter a number between 1 and {len(accounts)}{Style.RESET_ALL}")
            except ValueError:
                print(f"{Fore.RED}Please enter a valid number{Style.RESET_ALL}")

    def get_date_range(self) -> Tuple[datetime, datetime]:
        """
        Get date range from config file.
        
        Returns:
            Tuple[datetime, datetime]: Start and end dates.
        """
        date_config_file = os.path.join(config.CONFIG_DIR, "date_config.json")
        try:
            with open(date_config_file, "r") as f:
                date_config = json.load(f)
            
            start_date = datetime.strptime(date_config["start_date"], "%Y-%m-%d")
            end_date = datetime.strptime(date_config["end_date"], "%Y-%m-%d")
            
            return start_date, end_date
        except Exception as e:
            print(f"{Fore.RED}Error reading date configuration: {str(e)}{Style.RESET_ALL}")
            raise

    def configure_date_range(self):
        """Configure the date range in the config file."""
        date_config_file = os.path.join(config.CONFIG_DIR, "date_config.json")
        print(f"\n{Fore.YELLOW}Enter dates in format: YYYY-MM-DD{Style.RESET_ALL}")
        
        while True:
            try:
                start_str = input(f"{Fore.CYAN}Start date:{Style.RESET_ALL} ").strip()
                start_date = datetime.strptime(start_str, "%Y-%m-%d")
                break
            except ValueError:
                print(f"{Fore.RED}Invalid date format. Please use YYYY-MM-DD{Style.RESET_ALL}")

        while True:
            try:
                end_str = input(f"{Fore.CYAN}End date:{Style.RESET_ALL} ").strip()
                end_date = datetime.strptime(end_str, "%Y-%m-%d")
                
                if end_date < start_date:
                    print(f"{Fore.RED}End date must be after start date{Style.RESET_ALL}")
                    continue
                    
                break
            except ValueError:
                print(f"{Fore.RED}Invalid date format. Please use YYYY-MM-DD{Style.RESET_ALL}")

        config_data = {
            "start_date": start_date.strftime("%Y-%m-%d"),
            "end_date": end_date.strftime("%Y-%m-%d")
        }

        with open(date_config_file, "w") as f:
            json.dump(config_data, f, indent=4)
        
        print(f"{Fore.GREEN}Date range configuration saved successfully!{Style.RESET_ALL}")

    def confirm_action(self, message: str) -> bool:
        """
        Get user confirmation for an action.
        
        Args:
            message: Confirmation message to display.
            
        Returns:
            bool: True if confirmed, False otherwise.
        """
        response = input(f"\n{Fore.YELLOW}{message} (y/n):{Style.RESET_ALL} ").strip().lower()
        return response == "y"

    def start_operation(self, message: str):
        """
        Start a spinner for an operation.
        
        Args:
            message: Operation message to display.
        """
        self.spinner = Halo(text=message, spinner="dots")
        self.spinner.start()

    def stop_operation(self):
        """Stop the current operation spinner."""
        if self.spinner:
            self.spinner.stop()

    def display_error(self, message: str):
        """
        Display an error message.
        
        Args:
            message: Error message to display.
        """
        print(f"\n{Fore.RED}Error: {message}{Style.RESET_ALL}")

    def display_success(self, message: str):
        """
        Display a success message.
        
        Args:
            message: Success message to display.
        """
        print(f"\n{Fore.GREEN}{message}{Style.RESET_ALL}")

    def display_progress(self, total: int, desc: str = "Progress") -> tqdm:
        """
        Create and return a progress bar.
        
        Args:
            total: Total number of items.
            desc: Progress bar description.
            
        Returns:
            tqdm: Progress bar instance.
        """
        return tqdm(total=total, desc=desc, ncols=100) 