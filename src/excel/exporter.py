import logging
from datetime import datetime
import psycopg2
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from pathlib import Path
from typing import Dict, List
import concurrent.futures

class ExcelWriter:
    def __init__(self):
        self.workbook = None
        self.header_style = None
        self.cell_style = None
        self.setup_styles()

    def setup_styles(self):
        self.header_style = {
            'fill': PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid'),
            'font': Font(color='FFFFFF', bold=True),
            'border': Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            ),
            'alignment': Alignment(horizontal='center', vertical='center')
        }
        
        self.cell_style = {
            'border': Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        }

    def create_workbook(self):
        self.workbook = Workbook()
        self.workbook.remove(self.workbook.active)

    def write_dataframe(self, df: pd.DataFrame, sheet_name: str):
        sheet = self.workbook.create_sheet(title=sheet_name)
        
        # Write headers with style
        for col_idx, column in enumerate(df.columns, 1):
            cell = sheet.cell(row=1, column=col_idx, value=column)
            cell.fill = self.header_style['fill']
            cell.font = self.header_style['font']
            cell.border = self.header_style['border']
            cell.alignment = self.header_style['alignment']

        # Write data with style using openpyxl's cell API
        for row_idx, row in enumerate(df.values, 2):
            for col_idx, value in enumerate(row, 1):
                cell = sheet.cell(row=row_idx, column=col_idx, value=value)
                cell.border = self.cell_style['border']

        # Optimize column widths
        self._optimize_column_widths(sheet, df)

    def _optimize_column_widths(self, sheet, df: pd.DataFrame):
        for idx, column in enumerate(df.columns):
            max_length = max(
                df[column].astype(str).apply(len).max(),
                len(str(column))
            )
            sheet.column_dimensions[chr(65 + idx)].width = min(max_length + 2, 50)

    def save(self, filename: str):
        self.workbook.save(filename)

class DatabaseFetcher:
    def __init__(self, config: Dict):
        self.config = config
        self.conn = None

    def connect(self):
        try:
            self.conn = psycopg2.connect(**self.config)
            return True
        except Exception as e:
            logging.error(f"Connection error: {str(e)}")
            return False

    def fetch_view_data(self, view_name: str) -> pd.DataFrame:
        try:
            query = f'SELECT * FROM "{view_name}"'
            df = pd.read_sql_query(query, self.conn)
            return df
        except Exception as e:
            logging.error(f"Error fetching view {view_name}: {str(e)}")
            return pd.DataFrame()

    def close(self):
        if self.conn:
            self.conn.close()

def export_to_excel(environment_config: Dict, output_path: str = None) -> str:
    """
    Export database views to Excel file
    
    Args:
        environment_config: Dictionary containing database configuration and views
        output_path: Optional path for the output file. If None, generates a default path
    
    Returns:
        str: Path to the generated Excel file
    """
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = str(Path.cwd() / f"export_{timestamp}.xlsx")

    excel_writer = ExcelWriter()
    excel_writer.create_workbook()

    for db_config in environment_config:
        fetcher = DatabaseFetcher(db_config["config"])
        if not fetcher.connect():
            continue

        try:
            for view_name in db_config["views"]:
                df = fetcher.fetch_view_data(view_name)
                if not df.empty:
                    excel_writer.write_dataframe(df, view_name)
        finally:
            fetcher.close()

    excel_writer.save(output_path)
    logging.info(f"Excel file saved to: {output_path}")
    return output_path
