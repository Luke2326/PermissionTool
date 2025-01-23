import logging
from datetime import datetime
import psycopg2
import pandas as pd
from openpyxl import Workbook
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
    CHUNK_SIZE = 10000  # Numero di righe per chunk
    MAX_WORKERS = 4     # Numero massimo di thread per la scrittura

    def __init__(self):
        self.workbook = None
        self.header_style = None
        self.cell_style = None
        self.setup_styles()
        self._lock = threading.Lock()

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

    def _write_chunk(self, sheet, start_row: int, chunk: pd.DataFrame):
        """Scrive un chunk di dati nel foglio Excel"""
        with self._lock:  # Protegge l'accesso concorrente al foglio
            for row_idx, row in enumerate(chunk.itertuples(), start_row):
                for col_idx, value in enumerate(row[1:], 1):
                    cell = sheet.cell(row=row_idx, column=col_idx, value=value)
                    cell.border = self.cell_style['border']

    def write_dataframe(self, df: pd.DataFrame, sheet_name: str):
        """Scrive il DataFrame nel foglio Excel usando il multithreading"""
        logging.info(f"Creazione foglio: {sheet_name}")
        sheet = self.workbook.create_sheet(title=sheet_name)
        
        # Scrive l'header
        for col_idx, column in enumerate(df.columns, 1):
            cell = sheet.cell(row=1, column=col_idx, value=column)
            cell.fill = self.header_style['fill']
            cell.font = self.header_style['font']
            cell.border = self.header_style['border']
            cell.alignment = self.header_style['alignment']

        # Divide il DataFrame in chunks e li scrive in parallelo
        total_rows = len(df)
        chunks = [df[i:i + self.CHUNK_SIZE] for i in range(0, total_rows, self.CHUNK_SIZE)]
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=self.MAX_WORKERS) as executor:
            futures = []
            for i, chunk in enumerate(chunks):
                start_row = i * self.CHUNK_SIZE + 2  # +2 per l'header
                futures.append(
                    executor.submit(self._write_chunk, sheet, start_row, chunk)
                )
            
            # Monitora il progresso
            with tqdm(total=len(futures), desc=f"Scrittura {sheet_name}", unit="chunk") as pbar:
                for future in concurrent.futures.as_completed(futures):
                    future.result()  # Aspetta il completamento e gestisce eventuali eccezioni
                    pbar.update(1)

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
    CHUNK_SIZE = 50000  # Dimensione del chunk per l'estrazione dati

    def __init__(self, config: Dict):
        self.config = config
        self.conn = None

    def connect(self) -> bool:
        try:
            logging.info(f"Connessione al database {self.config.get('database')} su {self.config.get('host')}...")
            self.conn = psycopg2.connect(**self.config)
            self.conn.set_session(readonly=True)  # Ottimizzazione per query di sola lettura
            logging.info("Connessione al database stabilita con successo")
            return True
        except Exception as e:
            logging.error(f"Errore di connessione al database: {str(e)}")
            return False

    def _get_optimized_query(self, view_name: str) -> str:
        """Genera una query ottimizzata con paginazione"""
        return f'''
            SELECT *
            FROM "{view_name}"
            ORDER BY 1  -- Ordina per la prima colonna per consistenza
        '''

    def _stream_results(self, query: str) -> Generator[pd.DataFrame, None, None]:
        """Stream dei risultati in chunks"""
        try:
            for chunk in pd.read_sql_query(query, self.conn, chunksize=self.CHUNK_SIZE):
                yield chunk
        except Exception as e:
            logging.error(f"Errore nello streaming dei dati: {str(e)}")
            yield pd.DataFrame()

    def fetch_view_data(self, view_name: str) -> pd.DataFrame:
        """Estrae i dati dalla vista in modo ottimizzato"""
        try:
            logging.info(f"Inizio estrazione dati dalla vista: {view_name}")
            
            # Prima ottiene il conteggio totale
            count_query = f'SELECT COUNT(*) FROM "{view_name}"'
            total_rows = pd.read_sql_query(count_query, self.conn).iloc[0, 0]
            logging.info(f"Totale righe da estrarre da {view_name}: {total_rows}")

            # Se la vista è piccola, la estrae direttamente
            if total_rows < self.CHUNK_SIZE:
                query = self._get_optimized_query(view_name)
                df = pd.read_sql_query(query, self.conn)
                logging.info(f"Estratte {len(df)} righe dalla vista {view_name}")
                return df

            # Per viste grandi, usa lo streaming
            chunks = []
            query = self._get_optimized_query(view_name)
            
            with tqdm(total=total_rows, desc=f"Estrazione {view_name}", unit="righe") as pbar:
                for chunk in self._stream_results(query):
                    chunks.append(chunk)
                    pbar.update(len(chunk))

            df = pd.concat(chunks, ignore_index=True)
            logging.info(f"Completata estrazione di {len(df)} righe dalla vista {view_name}")
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
    excel_writer.save(output_path)
    logging.info(f"Export completato. File salvato in: {output_path}")
    return output_path
