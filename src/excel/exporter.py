import logging
from datetime import datetime
import psycopg2
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from pathlib import Path
from typing import Dict, List, Generator, Optional
import concurrent.futures
from tqdm import tqdm
import io
import threading
from queue import Queue
import time
import sys
import os

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)

# Tabelle che verranno gestite separatamente per performance
LARGE_TABLES = {'permissions', 'set_role_group_version'}

class ExcelWriter:
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
        self.main_file = None
        self.temp_dir = None
        self.large_files = {}

    def initialize(self, output_path: str):
        """Inizializza l'ambiente di scrittura"""
        self.main_file = output_path
        self.temp_dir = Path(output_path).parent / "temp_excel"
        self.temp_dir.mkdir(exist_ok=True)
        
        # Crea il file principale vuoto
        wb = Workbook()
        wb.save(self.main_file)

    def _apply_header_style(self, sheet):
        """Applica stile all'header"""
        for col in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=1, column=col)
            cell.fill = self.header_style['fill']
            cell.font = self.header_style['font']
            cell.border = self.header_style['border']
            cell.alignment = self.header_style['alignment']

    def _optimize_column_widths(self, sheet, sample_size=100):
        """Ottimizza larghezza colonne"""
        for col in range(1, sheet.max_column + 1):
            max_length = 0
            column = chr(64 + col)
            
            # Campiona solo alcune righe
            for row in range(1, min(sample_size, sheet.max_row + 1)):
                cell = sheet.cell(row=row, column=col)
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            
            sheet.column_dimensions[column].width = min(max_length + 2, 50)

    def write_dataframe(self, df: pd.DataFrame, sheet_name: str):
        """Scrive il DataFrame nel file appropriato"""
        start_time = time.time()
        
        if sheet_name.lower() in LARGE_TABLES:
            # Per tabelle grandi, scrivi in un file CSV separato
            csv_file = self.temp_dir / f"{sheet_name}.csv"
            logging.info(f"Scrittura {sheet_name} in CSV temporaneo...")
            df.to_csv(csv_file, index=False)
            self.large_files[sheet_name] = csv_file
            elapsed = time.time() - start_time
            logging.info(f"CSV {sheet_name} scritto in {elapsed:.2f} secondi")
        else:
            # Per tabelle piccole, scrivi direttamente nel file principale
            logging.info(f"Scrittura {sheet_name} nel file principale...")
            wb = load_workbook(self.main_file)
            
            if sheet_name in wb.sheetnames:
                wb.remove(wb[sheet_name])
            
            sheet = wb.create_sheet(sheet_name)
            
            # Scrivi header
            for col, name in enumerate(df.columns, 1):
                cell = sheet.cell(row=1, column=col, value=str(name))
            
            # Scrivi dati
            for row_idx, row in enumerate(df.values, 2):
                for col_idx, value in enumerate(row, 1):
                    sheet.cell(row=row_idx, column=col_idx, value=value)
            
            self._apply_header_style(sheet)
            self._optimize_column_widths(sheet)
            
            wb.save(self.main_file)
            elapsed = time.time() - start_time
            logging.info(f"Foglio {sheet_name} scritto in {elapsed:.2f} secondi")

    def finalize(self):
        """Unisce tutti i file in uno solo"""
        if not self.large_files:
            return

        logging.info("Unione dei file in corso...")
        start_time = time.time()

        # Crea un nuovo file Excel con pandas per le tabelle grandi
        large_file = self.temp_dir / "large_tables.xlsx"
        with pd.ExcelWriter(large_file, engine='openpyxl') as writer:
            for sheet_name, csv_file in self.large_files.items():
                logging.info(f"Processando {sheet_name}...")
                # Leggi il CSV in chunks per gestire la memoria
                chunks = pd.read_csv(csv_file, chunksize=50000)
                first_chunk = True
                for chunk in chunks:
                    if first_chunk:
                        chunk.to_excel(writer, sheet_name=sheet_name, index=False)
                        first_chunk = False
                    else:
                        chunk.to_excel(writer, sheet_name=sheet_name, startrow=writer.sheets[sheet_name].max_row + 1, header=False, index=False)

        # Copia i fogli dal file delle tabelle grandi al file principale
        logging.info("Copiando i fogli nel file principale...")
        large_wb = load_workbook(large_file)
        main_wb = load_workbook(self.main_file)

        for sheet_name in large_wb.sheetnames:
            if sheet_name in main_wb.sheetnames:
                main_wb.remove(main_wb[sheet_name])
            
            source_sheet = large_wb[sheet_name]
            new_sheet = main_wb.create_sheet(sheet_name)
            
            # Copia celle
            for row in source_sheet.rows:
                for cell in row:
                    new_sheet[cell.coordinate].value = cell.value
            
            self._apply_header_style(new_sheet)
            self._optimize_column_widths(new_sheet)

        main_wb.save(self.main_file)

        # Pulizia
        try:
            for csv_file in self.large_files.values():
                os.remove(csv_file)
            os.remove(large_file)
            os.rmdir(self.temp_dir)
        except:
            pass

        elapsed = time.time() - start_time
        logging.info(f"Unione completata in {elapsed:.2f} secondi")

class DatabaseFetcher:
    CHUNK_SIZE = 100000

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
    # Definisci il percorso base per le estrazioni
    base_path = Path(r"\\tsclient\V\Estrazioni")
    
    try:
        # Crea la directory se non esiste
        base_path.mkdir(parents=True, exist_ok=True)
        logging.info(f"Directory di output verificata: {base_path}")
    except Exception as e:
        logging.warning(f"Impossibile creare la directory {base_path}. Uso il percorso locale. Errore: {str(e)}")
        base_path = Path.cwd()

    if output_path is None:
        # Crea sottodirectory per anno e mese
        current_date = datetime.now()
        year_month = current_date.strftime("%Y_%m")
        year_month_path = base_path / year_month
        
        try:
            year_month_path.mkdir(parents=True, exist_ok=True)
            logging.info(f"Creata directory per anno/mese: {year_month_path}")
        except Exception as e:
            logging.warning(f"Impossibile creare la directory {year_month_path}. Uso il percorso base. Errore: {str(e)}")
            year_month_path = base_path

        timestamp = current_date.strftime("%Y%m%d_%H%M%S")
        output_path = str(year_month_path / f"export_{environment_name}_{timestamp}.xlsx")

    try:
        logging.info(f"Inizializzazione export per ambiente {environment_name}")
        excel_writer = ExcelWriter()
        excel_writer.initialize(output_path)

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
        excel_writer.finalize()
        logging.info(f"Export completato. File salvato in: {output_path}")
        return output_path
        
    except PermissionError as e:
        logging.error(f"Errore di permessi durante il salvataggio in {output_path}. Assicurarsi di avere i permessi necessari.")
        raise
    except Exception as e:
        logging.error(f"Errore durante l'export: {str(e)}")
        raise
