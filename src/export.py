"""
Export operations for the Gmail Export Tool.
"""
import logging
from typing import List, Dict
import pandas as pd
from datetime import datetime
from email.utils import parsedate_to_datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class ExportManager:
    """Handles export operations for email data."""

    def __init__(self):
        """Initialize the export manager."""
        self.logger = logging.getLogger(__name__)

    def export_to_excel(self, emails_data: list, output_file: str, email: str) -> bool:
        """
        Export email data to Excel file.
        
        Args:
            emails_data: List of email data dictionaries.
            output_file: Path to output Excel file.
            email: Email address being exported.
            
        Returns:
            bool: True if successful, False otherwise.
        """
        try:
            if not emails_data:
                self.logger.error("No data to export")
                return False

            # Convert to DataFrame
            df = pd.DataFrame(emails_data)
            
            # Convert RFC 2822 dates to datetime objects
            df["Date"] = pd.to_datetime(df["Date"])
            
            # Sort by date
            df = df.sort_values("Date")
            
            # Format date for display
            df["Date"] = df["Date"].dt.strftime("%Y-%m-%d %H:%M:%S")
            
            # Reorder columns
            columns = ["Date", "Username", "Domain", "Subject"]
            df = df[columns]
            
            # Export to Excel with formatting
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name=f"Sent Emails - {email}")
                
                # Get the workbook and worksheet
                workbook = writer.book
                worksheet = writer.sheets[f"Sent Emails - {email}"]
                
                # Format headers
                header_style = {
                    'font': Font(bold=True, color="FFFFFF"),
                    'fill': PatternFill(start_color="366092", end_color="366092", fill_type="solid"),
                    'alignment': Alignment(horizontal="center", vertical="center"),
                    'border': Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
                }
                
                for col in range(1, len(columns) + 1):
                    cell = worksheet.cell(row=1, column=col)
                    cell.font = header_style['font']
                    cell.fill = header_style['fill']
                    cell.alignment = header_style['alignment']
                    cell.border = header_style['border']
                
                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column = [cell for cell in column]
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width
            
            self.logger.info(f"Successfully exported {len(emails_data)} emails to {output_file}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error exporting to Excel: {str(e)}")
            return False 