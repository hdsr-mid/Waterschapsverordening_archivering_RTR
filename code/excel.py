"""
Excel handler for RTR archiving output.

This module handles the creation and formatting of Excel workbooks that contain
RTR activity data, including color-coded cells based on modification dates.
"""

import os
from datetime import datetime
import xlsxwriter

HEADER_INDICES_BEFORE_WERKINGSGEBIEDEN = 8
AANTAL_WERKZAAMHEDEN_COLUMN_INDEX = 2

class ExcelHandler:
    """Handles Excel workbook creation and data writing for RTR archiving."""
    
    def __init__(self, bestuursorgaan, base_dir, env, date, headers):
        """
        Initialize the Excel handler and create a new workbook.
        
        Args:
            bestuursorgaan (str): Name of the government body
            base_dir (str): Base directory for output files
            env (str): Environment ('prod' or 'pre')
            date (str): Date in DD-MM-YYYY format
            headers (list): Column headers for the Excel sheet
        """
        self.headers = headers
        self.workbook_path = self.generate_file_path(date, bestuursorgaan, env, "status", base_dir, extension="xlsx")
        self.workbook = xlsxwriter.Workbook(self.workbook_path)
        self.worksheet = self.workbook.add_worksheet()
        self.cell_format = self.create_format('white', bold=False, text_wrap=False, border=True)
        self.blue_format = self.create_format('#538DD5', bold=False, text_wrap=False, border=True)
        self.setup_worksheet()

    def create_format(self, color, bold, text_wrap, border):
        """Create a cell format with specified styling."""
        return self.workbook.add_format({
            'bg_color': color,
            'text_wrap': text_wrap,
            'align': 'left',
            'valign': 'top',
            'bold': bold,
            'border': border,
        })

    def setup_worksheet(self):
        """Set up the worksheet with headers, column widths, and frozen panes."""
        header_format = self.create_format('#DDDDDD', bold=True, text_wrap=False, border=True)
        self.worksheet.write_row('A1', self.headers, header_format)
        self.adjust_column_widths()
        self.worksheet.freeze_panes(1, 1)

    def adjust_column_widths(self):
        """Adjust column widths based on header lengths."""
        for i, header in enumerate(self.headers, 1):
            padding = 4
            column_width = 4 if i > HEADER_INDICES_BEFORE_WERKINGSGEBIEDEN else len(header) + padding
            self.worksheet.set_column(i - 1, i - 1, column_width)

    def write_data_to_cells(self, row, data_to_write):
        """
        Write a row of data to the worksheet with appropriate formatting.
        
        Args:
            row (int): Row number to write to
            data_to_write (list): Data values to write
        """
        for col, content in enumerate(data_to_write):
            if content == 1 and col != AANTAL_WERKZAAMHEDEN_COLUMN_INDEX:
                self.worksheet.write(row - 1, col, " ", self.blue_format)
            else:
                self.write_content(row - 1, col, content)

    def write_content(self, row, col, content):
        """Write content to a cell with color-coding based on date if applicable."""
        try:
            content_date = datetime.strptime(str(content), "%d-%m-%Y %H:%M:%S")
            color = self.determine_color_based_on_date(content_date)
            cell_format = self.create_format(color, bold=False, text_wrap=False, border=True)
            self.worksheet.write(row, col, content, cell_format)
        except ValueError:
            self.worksheet.write(row, col, content, self.cell_format)

    def determine_color_based_on_date(self, content_date):
        """Determine cell color based on how recent the date is."""
        difference = datetime.now() - content_date
        return self.set_green_intensity(difference.days)

    @staticmethod
    def set_green_intensity(days_diff):
        """
        Set green color intensity based on days difference.
        
        More recent dates get a more intense green color.
        
        Args:
            days_diff (int): Number of days since the date
            
        Returns:
            str: Hex color code
        """
        if days_diff < 1:
            return '#00FF00'
        elif days_diff < 8:
            return '#32CD32'
        elif days_diff < 30:
            return '#98FB98'
        elif days_diff < 60:
            return '#90EE90'
        else:
            return '#F0FFF0'

    def close_workbook(self):
        """Close and save the workbook."""
        self.workbook.close()

    @staticmethod
    def generate_file_path(date_str, overheid, env, suffix, base_dir, extension=None):
        """
        Generate a file path for output files.
        
        Args:
            date_str (str): Date in DD-MM-YYYY format
            overheid (str): Government body name
            env (str): Environment ('prod' or 'pre')
            suffix (str): File name suffix
            base_dir (str): Base directory
            extension (str, optional): File extension. If None, creates a directory path.
            
        Returns:
            str: Generated file path
        """
        date_from_arg = datetime.strptime(date_str, "%d-%m-%Y")
        formatted_date = date_from_arg.strftime("%Y%m%d")
        environment = "productie-omgeving" if env == "prod" else "pre-omgeving"
        base_name = f"{formatted_date}_{overheid.replace(' ', '_')}_{environment}_STTR_{suffix}"
        
        if extension:
            full_name = f"{base_name}.{extension}"
        else:
            full_name = base_name  # No extension, treat as folder
        
        return os.path.join(base_dir, f"log/{full_name}")
