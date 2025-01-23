import logging
from datetime import datetime
import psycopg2
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from pathlib import Path
from typing import Dict, List
import concurrent.futures
from tqdm import tqdm

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
        logging.info(f"Creazione foglio: {sheet_name}")
        sheet = self.workbook.create_sheet(title=sheet_name)
        
        # Write headers with style
        for col_idx, column in enumerate(df.columns, 1):
            cell = sheet.cell(row=1, column=col_idx, value=column)
            cell.fill = self.header_style['fill']
            cell.font = self.header_style['font']
            cell.border = self.header_style['border']
            cell.alignment = self.header_style['alignment']
        
        # Write data with progress bar
        total_rows = len(df)
        logging.info(f"Scrittura {total_rows} righe nel foglio {sheet_name}")
        
        for row_idx, row in enumerate(df.itertuples(), 2):
            for col_idx, value in enumerate(row[1:], 1):
                cell = sheet.cell(row=row_idx, column=col_idx, value=value)
                cell.border = self.cell_style['border']
        
        self._optimize_column_widths(sheet, df)
        logging.info(f"Completata la scrittura del foglio: {sheet_name}")

    def _optimize_column_widths(self, sheet, df: pd.DataFrame):
        for col_idx, column in enumerate(df.columns, 1):
            max_length = max(
                len(str(df[column].astype(str).max())),
                len(str(column))
            )
            adjusted_width = min(max_length + 2, 50)  # Cap width at 50
            sheet.column_dimensions[chr(64 + col_idx)].width = adjusted_width

    def save(self, filename: str):
        logging.info(f"Salvataggio del file Excel: {filename}")
        self.workbook.save(filename)
        logging.info("File Excel salvato con successo")

class DatabaseFetcher:
    def __init__(self, config: Dict):
        self.config = config
        self.conn = None

    def connect(self) -> bool:
        try:
            logging.info(f"Connessione al database {self.config.get('database')} su {self.config.get('host')}...")
            self.conn = psycopg2.connect(**self.config)
            logging.info("Connessione al database stabilita con successo")
            return True
        except Exception as e:
            logging.error(f"Errore di connessione al database: {str(e)}")
            return False

    def fetch_view_data(self, view_name: str) -> pd.DataFrame:
        try:
            logging.info(f"Estrazione dati dalla vista: {view_name}")
            query = f'SELECT * FROM "{view_name}"'
            df = pd.read_sql_query(query, self.conn)
            logging.info(f"Estratte {len(df)} righe dalla vista {view_name}")
            return df
        except Exception as e:
            logging.error(f"Errore nell'estrazione dalla vista {view_name}: {str(e)}")
            return pd.DataFrame()

    def close(self):
        if self.conn:
            self.conn.close()
            logging.info("Connessione al database chiusa")

def export_to_excel(environment_config: Dict, output_path: str = None, environment_name: str = "UNKNOWN") -> str:
    """
    Export database views to Excel file
    
    Args:
        environment_config: Dictionary containing database configuration and views
        output_path: Optional path for the output file. If None, generates a default path
        environment_name: Name of the environment (e.g., SIT, UAT, PREPROD)
    
    Returns:
        str: Path to the generated Excel file
    """
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = str(Path.cwd() / f"export_{environment_name}_{timestamp}.xlsx")

    logging.info(f"Inizializzazione export per ambiente {environment_name}")
    excel_writer = ExcelWriter()
    excel_writer.create_workbook()

    total_views = sum(len(db_config["views"]) for db_config in environment_config)
    progress_bar = tqdm(total=total_views, desc="Progresso estrazione", unit="vista")

    for db_config in environment_config:
        fetcher = DatabaseFetcher(db_config["config"])
        if not fetcher.connect():
            progress_bar.update(len(db_config["views"]))  # Skip views for failed connection
            continue

        try:
            for view_name in db_config["views"]:
                df = fetcher.fetch_view_data(view_name)
                if not df.empty:
                    excel_writer.write_dataframe(df, view_name)
                progress_bar.update(1)
                progress_bar.set_description(f"Elaborazione {view_name}")
        finally:
            fetcher.close()

    progress_bar.close()
    excel_writer.save(output_path)
    logging.info(f"Export completato. File salvato in: {output_path}")
    return output_path
