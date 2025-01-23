import logging
from datetime import datetime
import psycopg2
import pandas as pd
import pyarrow as pa
import pyarrow.csv as csv
import xlsxwriter
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
import gc  # Importa il modulo gc per gestione memoria
import tempfile  # Aggiungendo l'import mancante di tempfile

# Configura il logging per essere sempre visibile
class TqdmToLogger(io.StringIO):
    logger = None
    level = None
    buf = ''
    def __init__(self, logger, level=None):
        super(TqdmToLogger, self).__init__()
        self.logger = logger
        self.level = level or logging.INFO
    
    def write(self, buf):
        self.buf = buf.strip('\r\n\t ')
        if self.buf:
            self.logger.log(self.level, self.buf)
    
    def flush(self):
        if self.buf:
            self.logger.log(self.level, self.buf)
        self.buf = ''

# Configura il logger per scrivere su stdout con flush immediato
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger()
tqdm_logger = TqdmToLogger(logger, level=logging.INFO)

def log_info(message: str):
    """Wrapper per logging che forza il flush"""
    logger.info(message)
    sys.stdout.flush()

# Tabelle che verranno gestite separatamente per performance
LARGE_TABLES = {'permissions', 'set_role_group_version'}

class Exporter:
    """
    Classe ottimizzata per l'export di grandi quantità di dati in Excel.
    Usa pyarrow per la gestione efficiente dei dati e xlsxwriter per Excel.
    """
    
    def __init__(self, output_file: str):
        """
        Inizializza l'exporter
        
        Args:
            output_file: Percorso del file Excel di output
        """
        self.output_file = Path(output_file)
        self.temp_dir = Path(tempfile.mkdtemp())
        self.workbook = None
        self.formats = {}
        self.large_files = {}
        
        # Crea il workbook
        self.workbook = xlsxwriter.Workbook(
            self.output_file,
            {
                'constant_memory': True,
                'default_date_format': 'yyyy-mm-dd',
                'remove_timezone': True
            }
        )
        
        # Definisci i formati
        self._init_formats()
    
    def _init_formats(self):
        """Inizializza i formati Excel"""
        # Formato header: sfondo rosso, testo bianco, bordi
        self.formats['header'] = self.workbook.add_format({
            'bold': True,
            'bg_color': '#FF0000',  # Rosso
            'font_color': '#FFFFFF',  # Testo bianco
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Formato base per tutte le celle: bordi
        self.formats['base'] = self.workbook.add_format({
            'border': 1,
            'valign': 'top'
        })
        
        # Formato date con bordi
        self.formats['date'] = self.workbook.add_format({
            'num_format': 'yyyy-mm-dd',
            'border': 1,
            'valign': 'top'
        })

    def _adjust_column_width(self, worksheet, col_idx, header, data_sample):
        """
        Calcola la larghezza ottimale per una colonna basandosi sull'header e un campione di dati
        """
        # Lunghezza dell'header
        max_width = len(str(header)) + 2
        
        # Analizza il campione di dati
        for value in data_sample:
            if value is not None and pd.notna(value):
                # Gestisci i valori multi-riga
                lines = str(value).split('\n')
                for line in lines:
                    max_width = max(max_width, len(str(line)) + 2)
        
        # Limita la larghezza massima a 100 caratteri
        max_width = min(max_width, 100)
        worksheet.set_column(col_idx, col_idx, max_width)

    def write_dataframe(self, df: pd.DataFrame, sheet_name: str):
        """Scrive un DataFrame in un foglio Excel"""
        start_time = time.time()
        total_rows = len(df)
        
        try:
            if sheet_name.lower() in LARGE_TABLES:
                # Per tabelle grandi, usa pyarrow e file parquet
                log_info(f"Inizio scrittura {sheet_name} ({total_rows} righe) in formato parquet...")
                
                # Converti in tabella pyarrow
                table = pa.Table.from_pandas(df)
                
                # Scrivi in parquet
                parquet_file = self.temp_dir / f"{sheet_name}.parquet"
                pa.parquet.write_table(table, parquet_file)
                
                self.large_files[sheet_name] = parquet_file
                log_info(f"File parquet {sheet_name} completato")
                
                # Libera memoria
                del df, table
                gc.collect()
                
            else:
                # Per tabelle piccole, scrivi direttamente in Excel
                log_info(f"Inizio scrittura {sheet_name} ({total_rows} righe) nel file Excel...")
                
                # Crea il foglio
                worksheet = self.workbook.add_worksheet(sheet_name)
                
                # Scrivi header
                for col, name in enumerate(df.columns):
                    worksheet.write(0, col, str(name), self.formats['header'])
                
                # Calcola le larghezze delle colonne basandosi su un campione di dati
                sample_size = min(1000, len(df))
                for col, name in enumerate(df.columns):
                    sample_data = df.iloc[:sample_size, col].dropna().astype(str)
                    self._adjust_column_width(worksheet, col, name, sample_data)
                
                # Scrivi dati
                for row_idx, row in enumerate(df.itertuples(index=False), 1):
                    for col_idx, value in enumerate(row):
                        if pd.isna(value):
                            worksheet.write(row_idx, col_idx, '', self.formats['base'])
                        elif isinstance(value, pd.Timestamp):
                            worksheet.write_datetime(row_idx, col_idx, value.to_pydatetime(), self.formats['date'])
                        else:
                            worksheet.write(row_idx, col_idx, value, self.formats['base'])
                    
                    if row_idx % 10000 == 0:
                        log_info(f"Scritte {row_idx} righe di {total_rows} per {sheet_name}")
                        gc.collect()
                
                # Libera memoria
                del df
                gc.collect()
            
            elapsed = time.time() - start_time
            log_info(f"Elaborazione {sheet_name} completata in {elapsed:.2f} secondi")
            
        except Exception as e:
            log_info(f"Errore durante la scrittura di {sheet_name}: {str(e)}")
            raise

    def finalize(self):
        """Finalizza il file Excel unendo i file parquet"""
        if not self.large_files:
            self.workbook.close()
            return

        log_info("Inizio fase di unione dei file...")
        start_time = time.time()
        
        try:
            # Processa ogni file parquet
            for sheet_name, parquet_file in self.large_files.items():
                try:
                    log_info(f"Processando {sheet_name}...")
                    
                    # Crea il foglio
                    worksheet = self.workbook.add_worksheet(sheet_name)
                    
                    # Leggi il parquet in chunks usando pyarrow
                    reader = pa.parquet.ParquetFile(parquet_file)
                    schema = reader.schema
                    
                    # Scrivi header e inizializza array per le larghezze delle colonne
                    max_widths = [len(str(field)) + 2 for field in schema.names]
                    for col, field in enumerate(schema.names):
                        worksheet.write(0, col, str(field), self.formats['header'])
                    
                    # Scrivi dati in chunks
                    current_row = 1
                    for batch in reader.iter_batches(batch_size=50000):
                        df_chunk = batch.to_pandas()
                        
                        # Aggiorna le larghezze massime delle colonne
                        for col in range(len(schema.names)):
                            sample_data = df_chunk.iloc[:, col].dropna().astype(str)
                            for value in sample_data:
                                lines = str(value).split('\n')
                                for line in lines:
                                    max_widths[col] = max(max_widths[col], len(line) + 2)
                        
                        # Scrivi i dati
                        for row_idx, row in enumerate(df_chunk.itertuples(index=False), current_row):
                            for col_idx, value in enumerate(row):
                                if pd.isna(value):
                                    worksheet.write(row_idx, col_idx, '', self.formats['base'])
                                elif isinstance(value, pd.Timestamp):
                                    worksheet.write_datetime(row_idx, col_idx, value.to_pydatetime(), self.formats['date'])
                                else:
                                    worksheet.write(row_idx, col_idx, value, self.formats['base'])
                        
                        current_row += len(df_chunk)
                        
                        if current_row % 50000 == 0:
                            log_info(f"Processate {current_row-1} righe per {sheet_name}")
                            gc.collect()
                        
                        del df_chunk
                    
                    # Imposta le larghezze finali delle colonne
                    for col, width in enumerate(max_widths):
                        worksheet.set_column(col, col, min(width, 100))
                    
                    # Rimuovi il file parquet
                    os.remove(parquet_file)
                    log_info(f"File parquet {sheet_name} rimosso")
                    
                except Exception as e:
                    log_info(f"Errore durante l'elaborazione di {sheet_name}: {str(e)}")
                    raise

            # Chiudi il workbook
            self.workbook.close()
            
            # Pulizia directory temporanea
            try:
                os.rmdir(self.temp_dir)
                log_info("Directory temporanea rimossa")
            except Exception as e:
                log_info(f"Errore durante la rimozione della directory temporanea: {str(e)}")
            
            elapsed = time.time() - start_time
            log_info(f"Unione completata in {elapsed:.2f} secondi")
            
        except Exception as e:
            log_info(f"Errore durante la fase di unione: {str(e)}")
            raise
    
    def __del__(self):
        """Cleanup quando l'oggetto viene distrutto"""
        try:
            if self.workbook:
                self.workbook.close()
            
            # Rimuovi file temporanei
            if self.temp_dir.exists():
                for file in self.temp_dir.glob('*'):
                    try:
                        os.remove(file)
                    except:
                        pass
                try:
                    os.rmdir(self.temp_dir)
                except:
                    pass
        except:
            pass

class DatabaseFetcher:
    CHUNK_SIZE = 100000

    def __init__(self, config: Dict):
        self.config = config
        self.conn = None

    def connect(self) -> bool:
        try:
            log_info(f"Connessione al database {self.config.get('database')} su {self.config.get('host')}...")
            self.conn = psycopg2.connect(**self.config)
            self.conn.set_session(readonly=True)
            log_info("Connessione al database stabilita con successo")
            return True
        except Exception as e:
            log_info(f"Errore di connessione al database: {str(e)}")
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
            log_info(f"Inizio estrazione dati dalla vista: {view_name}")
            start_time = time.time()
            
            # Ottiene il conteggio totale
            count_query = f'SELECT COUNT(*) FROM "{view_name}"'
            total_rows = pd.read_sql_query(count_query, self.conn).iloc[0, 0]
            log_info(f"Totale righe da estrarre da {view_name}: {total_rows}")

            # Per viste piccole, estrazione diretta
            if total_rows < self.CHUNK_SIZE:
                query = self._get_optimized_query(view_name)
                df = pd.read_sql_query(query, self.conn)
                elapsed_time = time.time() - start_time
                log_info(f"Estratte {len(df)} righe da {view_name} in {elapsed_time:.2f} secondi")
                return df

            # Per viste grandi, usa COPY per massima performance
            buffer = io.StringIO()
            copy_sql = f'COPY (SELECT * FROM "{view_name}") TO STDOUT WITH CSV HEADER'
            
            with tqdm(total=total_rows, desc=f"Estrazione {view_name}", unit="righe", file=tqdm_logger) as pbar:
                cursor = self.conn.cursor()
                cursor.copy_expert(copy_sql, buffer)
                cursor.close()
                
                # Reset buffer e leggi con pandas
                buffer.seek(0)
                df = pd.read_csv(buffer)
                pbar.update(total_rows)

            elapsed_time = time.time() - start_time
            log_info(f"Estratte {len(df)} righe da {view_name} in {elapsed_time:.2f} secondi")
            return df

        except Exception as e:
            log_info(f"Errore nell'estrazione dalla vista {view_name}: {str(e)}")
            return pd.DataFrame()

    def close(self):
        if self.conn:
            self.conn.close()
            log_info("Connessione al database chiusa")

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
        log_info(f"Directory di output verificata: {base_path}")
    except Exception as e:
        log_info(f"Impossibile creare la directory {base_path}. Uso il percorso locale. Errore: {str(e)}")
        base_path = Path.cwd()

    if output_path is None:
        # Crea sottodirectory per anno e mese
        current_date = datetime.now()
        year_month = current_date.strftime("%Y_%m")
        year_month_path = base_path / year_month
        
        try:
            year_month_path.mkdir(parents=True, exist_ok=True)
            log_info(f"Creata directory per anno/mese: {year_month_path}")
        except Exception as e:
            log_info(f"Impossibile creare la directory {year_month_path}. Uso il percorso base. Errore: {str(e)}")
            year_month_path = base_path

        timestamp = current_date.strftime("%Y%m%d_%H%M%S")
        output_path = str(year_month_path / f"export_{environment_name}_{timestamp}.xlsx")

    try:
        log_info(f"Inizializzazione export per ambiente {environment_name}")
        exporter = Exporter(output_path)

        total_views = sum(len(db_config["views"]) for db_config in environment_config)
        progress_bar = tqdm(total=total_views, desc="Progresso totale", unit="vista", file=tqdm_logger)

        # Prima elabora tutte le tabelle piccole
        log_info("Fase 1: Elaborazione tabelle piccole")
        for db_config in environment_config:
            fetcher = DatabaseFetcher(db_config["config"])
            if not fetcher.connect():
                progress_bar.update(len(db_config["views"]))
                continue

            try:
                for view_name in db_config["views"]:
                    if view_name.lower() not in LARGE_TABLES:
                        start_time = time.time()
                        df = fetcher.fetch_view_data(view_name)
                        
                        if not df.empty:
                            exporter.write_dataframe(df, view_name)
                        
                        elapsed_time = time.time() - start_time
                        log_info(f"Elaborazione {view_name} completata in {elapsed_time:.2f} secondi")
                        
                        progress_bar.update(1)
                        progress_bar.set_description(f"Completata {view_name}")
                        
                        # Forza pulizia memoria
                        del df
                        gc.collect()
            finally:
                fetcher.close()

        # Poi elabora le tabelle grandi
        log_info("Fase 2: Elaborazione tabelle grandi")
        for db_config in environment_config:
            fetcher = DatabaseFetcher(db_config["config"])
            if not fetcher.connect():
                progress_bar.update(len(db_config["views"]))
                continue

            try:
                for view_name in db_config["views"]:
                    if view_name.lower() in LARGE_TABLES:
                        start_time = time.time()
                        df = fetcher.fetch_view_data(view_name)
                        
                        if not df.empty:
                            exporter.write_dataframe(df, view_name)
                        
                        elapsed_time = time.time() - start_time
                        log_info(f"Elaborazione {view_name} completata in {elapsed_time:.2f} secondi")
                        
                        progress_bar.update(1)
                        progress_bar.set_description(f"Completata {view_name}")
                        
                        # Forza pulizia memoria
                        del df
                        gc.collect()
            finally:
                fetcher.close()

        progress_bar.close()
        
        # Fase finale: unione dei file
        log_info("Fase 3: Unione dei file")
        exporter.finalize()
        log_info(f"Export completato. File salvato in: {output_path}")
        return output_path
        
    except PermissionError as e:
        log_info(f"Errore di permessi durante il salvataggio in {output_path}. Assicurarsi di avere i permessi necessari.")
        raise
    except Exception as e:
        log_info(f"Errore durante l'export: {str(e)}")
        raise
