import logging
from datetime import datetime
import psycopg2
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from pathlib import Path
from typing import Dict, List, Generator, Optional
import concurrent.futures
from tqdm import tqdm
import io
import threading
from queue import Queue
import time

class ExcelWriter:
    CHUNK_SIZE = 100000  # Aumentato il chunk size per ridurre le operazioni di I/O

    def __init__(self):
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
        self.output_file = None

    def create_excel(self, output_path: str):
        """Inizializza il file Excel"""
        self.output_file = output_path
        # Crea un Excel vuoto con il writer di pandas
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            pd.DataFrame().to_excel(writer)

    def _apply_styles(self, sheet_name: str):
        """Applica gli stili dopo la scrittura dei dati"""
        wb = load_workbook(self.output_file)
        sheet = wb[sheet_name]

        # Applica stili all'header
        for col in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=1, column=col)
            cell.fill = self.header_style['fill']
            cell.font = self.header_style['font']
            cell.border = self.header_style['border']
            cell.alignment = self.header_style['alignment']

        # Ottimizza larghezze colonne
        for col in range(1, sheet.max_column + 1):
            max_length = 0
            column = chr(64 + col)
            
            # Campiona solo alcune righe per la larghezza
            for row in range(1, min(1000, sheet.max_row + 1)):
                cell = sheet.cell(row=row, column=col)
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            
            adjusted_width = min(max_length + 2, 50)
            sheet.column_dimensions[column].width = adjusted_width

        wb.save(self.output_file)

    def write_dataframe(self, df: pd.DataFrame, sheet_name: str):
        """Scrive il DataFrame nel foglio Excel usando pandas"""
        logging.info(f"Inizio scrittura foglio: {sheet_name}")
        start_time = time.time()

        # Usa pandas per scrivere i dati in modo efficiente
        with pd.ExcelWriter(self.output_file, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Applica gli stili dopo la scrittura
        self._apply_styles(sheet_name)

        elapsed_time = time.time() - start_time
        logging.info(f"Scrittura foglio {sheet_name} completata in {elapsed_time:.2f} secondi")

class DatabaseFetcher:
    CHUNK_SIZE = 100000  # Aumentato anche qui per ridurre le operazioni di I/O

    def __init__(self, config: Dict):
        self.config = config
        self.conn = None

    def connect(self) -> bool:
        try:
            logging.info(f"Connessione al database {self.config.get('database')} su {self.config.get('host')}...")
            self.conn = psycopg2.connect(**self.config)
            self.conn.set_session(readonly=True)
            logging.info("Connessione al database stabilita con successo")
            return True
        except Exception as e:
            logging.error(f"Errore di connessione al database: {str(e)}")
            return False

    def _get_optimized_query(self, view_name: str) -> str:
        """Genera una query ottimizzata"""
        return f'''
            SELECT *
            FROM "{view_name}"
            ORDER BY 1
        '''

    def fetch_view_data(self, view_name: str) -> pd.DataFrame:
        """Estrae i dati dalla vista in modo ottimizzato"""
        try:
            logging.info(f"Inizio estrazione dati dalla vista: {view_name}")
            start_time = time.time()
            
            # Ottiene il conteggio totale
            count_query = f'SELECT COUNT(*) FROM "{view_name}"'
            total_rows = pd.read_sql_query(count_query, self.conn).iloc[0, 0]
            logging.info(f"Totale righe da estrarre da {view_name}: {total_rows}")

            # Per viste piccole, estrazione diretta
            if total_rows < self.CHUNK_SIZE:
                query = self._get_optimized_query(view_name)
                df = pd.read_sql_query(query, self.conn)
                elapsed_time = time.time() - start_time
                logging.info(f"Estratte {len(df)} righe da {view_name} in {elapsed_time:.2f} secondi")
                return df

            # Per viste grandi, usa COPY per massima performance
            buffer = io.StringIO()
            copy_sql = f'COPY (SELECT * FROM "{view_name}") TO STDOUT WITH CSV HEADER'
            
            with tqdm(total=total_rows, desc=f"Estrazione {view_name}", unit="righe") as pbar:
                cursor = self.conn.cursor()
                cursor.copy_expert(copy_sql, buffer)
                cursor.close()
                
                # Reset buffer e leggi con pandas
                buffer.seek(0)
                df = pd.read_csv(buffer)
                pbar.update(total_rows)

            elapsed_time = time.time() - start_time
            logging.info(f"Estratte {len(df)} righe da {view_name} in {elapsed_time:.2f} secondi")
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
    excel_writer.create_excel(output_path)

    total_views = sum(len(db_config["views"]) for db_config in environment_config)
    progress_bar = tqdm(total=total_views, desc="Progresso totale", unit="vista")

    for db_config in environment_config:
        fetcher = DatabaseFetcher(db_config["config"])
        if not fetcher.connect():
            progress_bar.update(len(db_config["views"]))
            continue

        try:
            for view_name in db_config["views"]:
                start_time = time.time()
                df = fetcher.fetch_view_data(view_name)
                
                if not df.empty:
                    excel_writer.write_dataframe(df, view_name)
                
                elapsed_time = time.time() - start_time
                logging.info(f"Elaborazione {view_name} completata in {elapsed_time:.2f} secondi")
                
                progress_bar.update(1)
                progress_bar.set_description(f"Completata {view_name}")
        finally:
            fetcher.close()

    progress_bar.close()
    logging.info(f"Export completato. File salvato in: {output_path}")
    return output_path
