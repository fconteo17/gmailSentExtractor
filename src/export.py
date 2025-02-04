"""
Export operations for the Gmail Export Tool.
"""
import logging
from typing import List, Dict
import pandas as pd
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class ExportManager:
    """Handles all export operations."""

    def __init__(self):
        """Initialize the ExportManager."""
        self.logger = logging.getLogger(__name__)

    def export_to_excel(
        self,
        emails_data: List[Dict],
        output_file: str,
        email_address: str
    ) -> bool:
        """
        Export email data to Excel file with formatting.
        
        Args:
            emails_data: List of email data dictionaries.
            output_file: Path to output Excel file.
            email_address: Email address associated with the export.
            
        Returns:
            bool: True if successful, False otherwise.
        """
        try:
            if not emails_data:
                self.logger.warning("No data to export")
                return False

            # Create DataFrame
            df = pd.DataFrame(emails_data)
            
            # Convert date strings to datetime objects for sorting
            df["Date"] = pd.to_datetime(df["Date"], format="mixed")
            
            # Sort by date
            df = df.sort_values("Date", ascending=False)
            
            # Format date column
            df["Date"] = df["Date"].dt.strftime("%Y-%m-%d %H:%M:%S")

            # Create Excel writer
            with pd.ExcelWriter(
                output_file,
                engine="openpyxl",
                datetime_format="yyyy-mm-dd hh:mm:ss"
            ) as writer:
                # Write data
                df.to_excel(writer, index=False, sheet_name="Sent Emails")
                
                # Get workbook and worksheet
                workbook = writer.book
                worksheet = writer.sheets["Sent Emails"]

                # Define styles
                header_font = Font(bold=True, color="FFFFFF", size=11)
                header_fill = PatternFill(
                    start_color="366092",
                    end_color="366092",
                    fill_type="solid"
                )
                cell_border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                
                # Format headers
                for col in range(1, len(df.columns) + 1):
                    cell = worksheet.cell(row=1, column=col)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = cell_border

                # Format data cells and adjust column widths
                for col in range(1, len(df.columns) + 1):
                    column_letter = get_column_letter(col)
                    max_length = 0
                    column = df.iloc[:, col-1]
                    
                    # Apply borders and center alignment to all cells in column
                    for row in range(2, len(df) + 2):  # +2 because Excel is 1-based and we have header
                        cell = worksheet.cell(row=row, column=col)
                        cell.border = cell_border
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                        
                        # Get maximum length for column width
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    
                    # Account for column header length
                    header_length = len(str(df.columns[col-1]))
                    max_length = max(max_length, header_length)
                    
                    # Set width with some padding
                    adjusted_width = (max_length + 2) * 1.2
                    worksheet.column_dimensions[column_letter].width = adjusted_width

                # Add metadata
                metadata_row = len(df) + 3
                metadata_font = Font(bold=True)
                
                # Add metadata with improved formatting
                metadata_cells = [
                    ("Export Information", True),
                    (f"Email Account: {email_address}", False),
                    (f"Export Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", False),
                    (f"Total Emails: {len(df)}", False)
                ]
                
                for i, (text, is_header) in enumerate(metadata_cells):
                    cell = worksheet.cell(row=metadata_row + i, column=1)
                    cell.value = text
                    if is_header:
                        cell.font = metadata_font

            self.logger.info(f"Successfully exported {len(df)} emails to {output_file}")
            return True

        except Exception as e:
            self.logger.error(f"Error exporting to Excel: {str(e)}")
            return False 