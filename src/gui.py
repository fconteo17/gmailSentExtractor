import sys
from datetime import datetime, date
import threading
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QLineEdit,
    QListWidget,
    QFrame,
    QMessageBox,
    QStatusBar,
    QCalendarWidget,
    QDialog,
    QProgressBar,
    QCheckBox,
    QListWidgetItem,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QDate
from PyQt6.QtGui import QFont
from calendar import monthrange
from tqdm import tqdm

from src.main import GmailExportTool


class CustomTqdm:
    def __init__(self, total, desc, callback):
        self.total = total
        self.n = 0
        self.desc = desc
        self.callback = callback
        if self.callback:
            self.callback(0, self.total)  # Initial progress

    def update(self, n):
        self.n += n
        if self.callback:
            self.callback(self.n, self.total)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass  # Cleanup if needed

    def close(self):
        pass

    def clear(self):
        pass

    def refresh(self):
        pass

    def write(self, s):
        pass

    @property
    def format_dict(self):
        return {"n": self.n, "total": self.total}


class EmailExportWorker(QThread):
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(str)
    progress_update = pyqtSignal(int, int)  # current, total

    def __init__(self, gmail_tool, email, start_date, end_date):
        super().__init__()
        self.gmail_tool = gmail_tool
        self.email = email
        self.start_date = start_date
        self.end_date = end_date

    def run(self):
        try:
            service = self.gmail_tool.gmail_service.setup_service(self.email)
            if not service:
                self.finished.emit(False, "Failed to authenticate with Gmail")
                return

            self.progress.emit("Fetching emails...")
            
            # Store original tqdm
            original_tqdm = tqdm
            
            # Create our tqdm replacement
            def custom_tqdm(*args, **kwargs):
                return CustomTqdm(
                    total=kwargs.get('total', args[0] if args else None),
                    desc=kwargs.get('desc', ''),
                    callback=self.progress_update.emit
                )
            
            # Replace tqdm globally
            tqdm.__init__ = custom_tqdm
            tqdm.__call__ = custom_tqdm
            tqdm.__new__ = lambda cls, *args, **kwargs: custom_tqdm(*args, **kwargs)
            
            try:
                emails_data = self.gmail_tool.gmail_service.get_sent_emails(
                    service, 
                    self.start_date, 
                    self.end_date
                )
            finally:
                # Restore original tqdm
                tqdm.__init__ = original_tqdm.__init__
                tqdm.__call__ = original_tqdm.__call__
                tqdm.__new__ = original_tqdm.__new__

            if not emails_data:
                self.finished.emit(False, "No emails found in the specified date range")
                return

            self.progress.emit("Exporting to Excel...")
            output_file = self.gmail_tool.file_manager.get_export_path(
                self.email, self.start_date
            )
            if self.gmail_tool.export_manager.export_to_excel(
                emails_data, output_file, self.email
            ):
                self.finished.emit(
                    True,
                    f"Successfully exported {len(emails_data)} emails to:\n{output_file}",
                )
            else:
                self.finished.emit(False, "Failed to export emails to Excel")
        except Exception as e:
            self.finished.emit(False, str(e))


class AddAccountWorker(QThread):
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(str)

    def __init__(self, gmail_tool, email):
        super().__init__()
        self.gmail_tool = gmail_tool
        self.email = email

    def run(self):
        try:
            self.progress.emit("Adding account...")
            if not self.gmail_tool.account_manager.add_account(self.email):
                self.finished.emit(False, "Failed to add account")
                return

            self.progress.emit("Authenticating with Gmail...")
            # Test authentication immediately
            service = self.gmail_tool.gmail_service.setup_service(self.email)
            if service:
                self.finished.emit(True, "Account added and authenticated successfully")
            else:
                # If authentication fails, remove the account
                self.gmail_tool.account_manager.remove_account(self.email)
                self.finished.emit(False, "Failed to authenticate with Gmail")
        except Exception as e:
            # If any error occurs, ensure the account is removed
            try:
                self.gmail_tool.account_manager.remove_account(self.email)
            except:
                pass
            self.finished.emit(False, str(e))


class CalendarDialog(QDialog):
    def __init__(self, parent=None, title="Select Date"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)

        layout = QVBoxLayout(self)

        # Calendar widget
        self.calendar = QCalendarWidget()
        self.calendar.setStyleSheet(
            """
            QCalendarWidget {
                background-color: #2b2b2b;
                color: white;
            }
            QCalendarWidget QToolButton {
                color: white;
                background-color: #2b2b2b;
            }
            QCalendarWidget QMenu {
                background-color: #2b2b2b;
                color: white;
            }
            QCalendarWidget QTableView {
                background-color: #2b2b2b;
                selection-background-color: #0078d4;
            }
        """
        )
        layout.addWidget(self.calendar)

        # Buttons
        button_layout = QHBoxLayout()

        ok_button = QPushButton("OK")
        ok_button.setStyleSheet(
            """
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #1084d8;
            }
        """
        )
        ok_button.clicked.connect(self.accept)

        cancel_button = QPushButton("Cancel")
        cancel_button.setStyleSheet(
            """
            QPushButton {
                background-color: #333333;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #404040;
            }
        """
        )
        cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

        self.setFixedSize(400, 300)

    def get_selected_date(self):
        return self.calendar.selectedDate()


class AccountListItem(QWidget):
    def __init__(self, email, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(10)
        
        # Create a container for checkbox to control its size
        checkbox_container = QWidget()
        checkbox_container.setFixedWidth(24)  # Reduced width
        checkbox_layout = QHBoxLayout(checkbox_container)
        checkbox_layout.setContentsMargins(0, 0, 0, 0)
        checkbox_layout.setSpacing(0)
        
        self.checkbox = QCheckBox()
        self.checkbox.setStyleSheet("""
            QCheckBox {
                spacing: 0px;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
            }
            QCheckBox::indicator:unchecked {
                border: 1px solid #3f3f3f;
                background: #2b2b2b;
                border-radius: 3px;
            }
            QCheckBox::indicator:checked {
                border: 1px solid #0078d4;
                background: #0078d4;
                border-radius: 3px;
                image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='white' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='20 6 9 17 4 12'%3E%3C/polyline%3E%3C/svg%3E");
            }
            QCheckBox::indicator:hover {
                border-color: #0078d4;
            }
        """)
        checkbox_layout.addWidget(self.checkbox, 0, Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(checkbox_container)
        
        self.email_label = QLabel(email)
        self.email_label.setStyleSheet("""
            QLabel {
                color: white;
                font-size: 12px;
                padding: 2px;
            }
        """)
        layout.addWidget(self.email_label, stretch=1)

    def is_checked(self):
        return self.checkbox.isChecked()

    def get_email(self):
        return self.email_label.text()

    def set_checked(self, checked):
        self.checkbox.setChecked(checked)


class GmailExportGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.gmail_tool = GmailExportTool()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Gmail Export Tool")
        self.setFixedSize(800, 500)

        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)  # Changed to QVBoxLayout

        # Create content area
        content_widget = QWidget()
        content_layout = QHBoxLayout(content_widget)
        
        # Create sidebar
        sidebar = self.create_sidebar()
        content_layout.addWidget(sidebar)

        # Create main content
        main_content = self.create_main_content()
        content_layout.addWidget(main_content)

        # Set layout proportions
        content_layout.setStretch(0, 1)  # Sidebar
        content_layout.setStretch(1, 2)  # Main content

        main_layout.addWidget(content_widget)

        # Create footer
        footer = QFrame()
        footer.setFrameStyle(QFrame.Shape.Box)
        footer.setStyleSheet("""
            QFrame {
                background-color: #2b2b2b;
                border-top: 1px solid #3f3f3f;
                border-bottom: none;
                border-left: none;
                border-right: none;
                border-radius: 0;
            }
        """)
        footer_layout = QVBoxLayout(footer)
        footer_layout.setContentsMargins(10, 5, 10, 5)

        # Progress info layout
        progress_info = QHBoxLayout()
        
        # Status area (left side)
        self.status_bar = QStatusBar()
        self.status_bar.setStyleSheet("""
            QStatusBar {
                border: none;
                background: transparent;
                color: #cccccc;
            }
        """)
        progress_info.addWidget(self.status_bar, stretch=1)
        
        # Progress area (right side)
        progress_layout = QHBoxLayout()
        self.progress_label = QLabel("")
        self.progress_label.setStyleSheet("color: #cccccc;")
        progress_layout.addWidget(self.progress_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #3f3f3f;
                border-radius: 4px;
                background-color: #1e1e1e;
                color: white;
                text-align: center;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #0078d4;
                border-radius: 3px;
            }
        """)
        self.progress_bar.setFixedWidth(200)
        self.progress_bar.hide()
        progress_layout.addWidget(self.progress_bar)
        
        progress_info.addLayout(progress_layout)
        footer_layout.addLayout(progress_info)
        
        main_layout.addWidget(footer)

        # Refresh account list
        self.refresh_account_list()

    def get_last_month_range(self):
        today = date.today()
        
        # If we're in month M, we want month M-1
        if today.month == 1:  # If January, go to previous year's December
            start_date = date(today.year - 1, 12, 1)
            _, last_day = monthrange(today.year - 1, 12)
            end_date = date(today.year - 1, 12, last_day)
        else:
            start_date = date(today.year, today.month - 1, 1)
            _, last_day = monthrange(today.year, today.month - 1)
            end_date = date(today.year, today.month - 1, last_day)
        
        return start_date, end_date

    def create_sidebar(self):
        sidebar = QFrame()
        sidebar.setFrameStyle(QFrame.Shape.Box)
        layout = QVBoxLayout(sidebar)

        # Title
        title = QLabel("Gmail Export Tool")
        title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # Accounts label and select all checkbox
        accounts_header = QWidget()
        header_layout = QHBoxLayout(accounts_header)
        header_layout.setContentsMargins(0, 0, 0, 0)
        
        accounts_label = QLabel("Accounts")
        accounts_label.setFont(QFont("Arial", 11))
        header_layout.addWidget(accounts_label)
        
        self.select_all_checkbox = QCheckBox("Select All")
        self.select_all_checkbox.setStyleSheet("""
            QCheckBox {
                color: #cccccc;
                spacing: 5px;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
            }
            QCheckBox::indicator:unchecked {
                border: 1px solid #3f3f3f;
                background: #2b2b2b;
                border-radius: 3px;
            }
            QCheckBox::indicator:checked {
                border: 1px solid #0078d4;
                background: #0078d4;
                border-radius: 3px;
                image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='white' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='20 6 9 17 4 12'%3E%3C/polyline%3E%3C/svg%3E");
            }
            QCheckBox::indicator:hover {
                border-color: #0078d4;
            }
        """)
        self.select_all_checkbox.clicked.connect(self.toggle_all_accounts)
        header_layout.addWidget(self.select_all_checkbox, alignment=Qt.AlignmentFlag.AlignRight)
        
        layout.addWidget(accounts_header)

        # Account list
        self.account_list = QListWidget()
        self.account_list.setStyleSheet("""
            QListWidget {
                background-color: #2b2b2b;
                color: white;
                border: 1px solid #3f3f3f;
                border-radius: 4px;
            }
            QListWidget::item {
                padding: 0px;
                border-bottom: 1px solid #3f3f3f;
            }
            QListWidget::item:selected {
                background-color: #363636;
            }
            QListWidget::item:hover {
                background-color: #323232;
            }
        """)
        self.account_list.setSpacing(1)  # Add small spacing between items
        layout.addWidget(self.account_list)

        # Remove account button
        remove_button = QPushButton("Remove Account")
        remove_button.setStyleSheet(
            """
            QPushButton {
                background-color: #d42828;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 4px;
                margin-top: 5px;
            }
            QPushButton:hover {
                background-color: #e13131;
            }
            QPushButton:pressed {
                background-color: #c42424;
            }
            QPushButton:disabled {
                background-color: #666666;
            }
        """
        )
        remove_button.clicked.connect(self.remove_account)
        layout.addWidget(remove_button)

        # Email input
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("Enter email address")
        self.email_input.setStyleSheet(
            """
            QLineEdit {
                padding: 8px;
                border: 1px solid #3f3f3f;
                border-radius: 4px;
                background-color: #2b2b2b;
                color: white;
                margin-top: 10px;
            }
        """
        )
        layout.addWidget(self.email_input)

        # Add account button
        add_button = QPushButton("Add Account")
        add_button.setStyleSheet(
            """
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #1084d8;
            }
            QPushButton:pressed {
                background-color: #006cbd;
            }
        """
        )
        add_button.clicked.connect(self.add_account)
        layout.addWidget(add_button)

        layout.addStretch()
        return sidebar

    def create_main_content(self):
        content = QFrame()
        content.setFrameStyle(QFrame.Shape.Box)
        layout = QVBoxLayout(content)

        # Title
        title = QLabel("Export Settings")
        title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # Date selection
        date_frame = QFrame()
        date_layout = QHBoxLayout(date_frame)

        # Get default date range (last month)
        default_start_date, default_end_date = self.get_last_month_range()

        # Start date
        start_date_layout = QVBoxLayout()
        start_date_label = QLabel("Start Date:")
        self.start_date_button = QPushButton("Select Start Date")
        self.start_date_button.setStyleSheet(
            """
            QPushButton {
                background-color: #2b2b2b;
                color: white;
                border: 1px solid #3f3f3f;
                border-radius: 4px;
                padding: 8px;
                min-width: 150px;
                text-align: left;
            }
            QPushButton:hover {
                background-color: #363636;
            }
        """
        )
        self.start_date = QDate(default_start_date.year, default_start_date.month, default_start_date.day)
        self.start_date_button.setText(self.start_date.toString("yyyy-MM-dd"))
        self.start_date_button.clicked.connect(self.show_start_date_dialog)

        start_date_layout.addWidget(start_date_label)
        start_date_layout.addWidget(self.start_date_button)
        date_layout.addLayout(start_date_layout)

        # Add some spacing between date selectors
        date_layout.addSpacing(20)

        # End date
        end_date_layout = QVBoxLayout()
        end_date_label = QLabel("End Date:")
        self.end_date_button = QPushButton("Select End Date")
        self.end_date_button.setStyleSheet(
            """
            QPushButton {
                background-color: #2b2b2b;
                color: white;
                border: 1px solid #3f3f3f;
                border-radius: 4px;
                padding: 8px;
                min-width: 150px;
                text-align: left;
            }
            QPushButton:hover {
                background-color: #363636;
            }
        """
        )
        self.end_date = QDate(default_end_date.year, default_end_date.month, default_end_date.day)
        self.end_date_button.setText(self.end_date.toString("yyyy-MM-dd"))
        self.end_date_button.clicked.connect(self.show_end_date_dialog)

        end_date_layout.addWidget(end_date_label)
        end_date_layout.addWidget(self.end_date_button)
        date_layout.addLayout(end_date_layout)

        # Center the date selection frame
        date_layout.addStretch()
        date_layout.insertStretch(0)

        layout.addWidget(date_frame)
        layout.addSpacing(20)  # Add space before export button

        # Export button
        self.export_button = QPushButton("Export Emails")
        self.export_button.setStyleSheet(
            """
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 4px;
                font-weight: bold;
                min-width: 200px;
            }
            QPushButton:hover {
                background-color: #1084d8;
            }
            QPushButton:pressed {
                background-color: #006cbd;
            }
            QPushButton:disabled {
                background-color: #666666;
            }
        """
        )
        self.export_button.clicked.connect(self.export_emails)

        # Center the export button
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(self.export_button)
        button_layout.addStretch()
        layout.addLayout(button_layout)

        layout.addStretch()
        return content

    def refresh_account_list(self):
        self.account_list.clear()
        accounts = self.gmail_tool.account_manager.list_accounts()
        for email in accounts:
            item = QListWidgetItem(self.account_list)
            account_widget = AccountListItem(email)
            item.setSizeHint(account_widget.sizeHint())
            self.account_list.addItem(item)
            self.account_list.setItemWidget(item, account_widget)

    def toggle_all_accounts(self, checked):
        for i in range(self.account_list.count()):
            item = self.account_list.item(i)
            widget = self.account_list.itemWidget(item)
            widget.set_checked(checked)

    def get_selected_accounts(self):
        selected = []
        for i in range(self.account_list.count()):
            item = self.account_list.item(i)
            widget = self.account_list.itemWidget(item)
            if widget.is_checked():
                selected.append(widget.get_email())
        return selected

    def add_account(self):
        email = self.email_input.text().strip()
        if not email:
            QMessageBox.critical(self, "Error", "Please enter an email address")
            return

        # Disable the input and buttons during the process
        self.email_input.setEnabled(False)
        self.export_button.setEnabled(False)
        for button in self.findChildren(QPushButton):
            button.setEnabled(False)

        self.worker = AddAccountWorker(self.gmail_tool, email)
        self.worker.progress.connect(lambda msg: self.status_bar.showMessage(msg))
        self.worker.finished.connect(self.on_add_account_finished)
        self.worker.start()

    def on_add_account_finished(self, success, message):
        # Re-enable the input and buttons
        self.email_input.setEnabled(True)
        self.export_button.setEnabled(True)
        for button in self.findChildren(QPushButton):
            button.setEnabled(True)

        if success:
            self.status_bar.showMessage(message)
            self.email_input.clear()
            self.refresh_account_list()
            QMessageBox.information(self, "Success", message)
        else:
            QMessageBox.critical(self, "Error", message)
        self.status_bar.showMessage("")

    def show_start_date_dialog(self):
        dialog = CalendarDialog(self, "Select Start Date")
        dialog.calendar.setSelectedDate(self.start_date)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.start_date = dialog.get_selected_date()
            self.start_date_button.setText(self.start_date.toString("yyyy-MM-dd"))

    def show_end_date_dialog(self):
        dialog = CalendarDialog(self, "Select End Date")
        dialog.calendar.setSelectedDate(self.end_date)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.end_date = dialog.get_selected_date()
            self.end_date_button.setText(self.end_date.toString("yyyy-MM-dd"))

    def export_emails(self):
        selected_accounts = self.get_selected_accounts()
        if not selected_accounts:
            QMessageBox.critical(self, "Error", "Please select at least one account")
            return

        start_date = self.start_date.toPyDate()
        end_date = self.end_date.toPyDate()

        if start_date > end_date:
            QMessageBox.critical(self, "Error", "Start date must be before end date")
            return

        self.export_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        # Start with the first account
        self.current_account_index = 0
        self.selected_accounts = selected_accounts
        self.start_next_export()

    def start_next_export(self):
        if self.current_account_index >= len(self.selected_accounts):
            self.export_button.setEnabled(True)
            self.progress_bar.hide()
            self.progress_label.setText("")
            self.status_bar.showMessage("")
            QMessageBox.information(self, "Success", "All exports completed successfully!")
            return

        email = self.selected_accounts[self.current_account_index]
        self.status_bar.showMessage(f"Processing account: {email}")
        
        self.worker = EmailExportWorker(
            self.gmail_tool, 
            email, 
            self.start_date.toPyDate(), 
            self.end_date.toPyDate()
        )
        self.worker.progress.connect(lambda msg: self.status_bar.showMessage(f"{email}: {msg}"))
        self.worker.progress_update.connect(self.update_progress)
        self.worker.finished.connect(self.on_single_export_finished)
        self.worker.start()

    def on_single_export_finished(self, success, message):
        email = self.selected_accounts[self.current_account_index]
        if success:
            self.status_bar.showMessage(f"Export completed for {email}")
        else:
            QMessageBox.warning(
                self, 
                "Warning", 
                f"Failed to export emails for {email}:\n{message}"
            )
        
        # Move to next account
        self.current_account_index += 1
        self.start_next_export()

    def update_progress(self, current, total):
        if total > 0:
            self.progress_bar.setVisible(True)
            percentage = int((current / total) * 100)
            self.progress_bar.setValue(percentage)
            self.progress_label.setText(f"Processing: {current}/{total} emails")
            self.progress_bar.setFormat(f"{percentage}%")

    def remove_account(self):
        selected_accounts = self.get_selected_accounts()
        if not selected_accounts:
            QMessageBox.warning(self, "Warning", "Please select at least one account to remove")
            return
            
        accounts_str = "\n".join(selected_accounts)
        reply = QMessageBox.question(
            self,
            "Confirm Removal",
            f"Are you sure you want to remove these accounts:\n{accounts_str}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                for email in selected_accounts:
                    self.gmail_tool.account_manager.remove_account(email)
                self.status_bar.showMessage("Account(s) removed successfully")
                self.refresh_account_list()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to remove account(s): {str(e)}")
            finally:
                self.status_bar.showMessage("")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # Use Fusion style for a modern look

    # Set dark theme
    app.setStyleSheet(
        """
        QMainWindow, QWidget {
            background-color: #1e1e1e;
            color: white;
        }
        QFrame {
            border: 1px solid #3f3f3f;
            border-radius: 4px;
        }
        QLabel {
            border: none;
        }
    """
    )

    window = GmailExportGUI()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
