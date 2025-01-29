import logging
from datetime import datetime
import psycopg2
import pandas as pd
import pyarrow as pa
import pyarrow.csv as csv
import xlsxwriter
from pathlib import Path
from typing import Dict, List, Generator, Optional, Any
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
import tkinter as tk
from tkinter import filedialog

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
    
    CHUNK_SIZE = 100000  # Aumentato per migliori performance
    MAX_COLUMN_WIDTH = 100
    MAX_WORKERS = 4  # Numero di worker per il processing parallelo
    
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
        self.column_widths = {}
        self.executor = concurrent.futures.ThreadPoolExecutor(max_workers=self.MAX_WORKERS)
        self._is_closed = False
        
        # Opzioni ottimizzate per il workbook
        self.workbook = xlsxwriter.Workbook(
            self.output_file,
            {
                'constant_memory': True,
                'default_date_format': 'yyyy-mm-dd',
                'remove_timezone': True,
                'strings_to_numbers': True,
                'strings_to_formulas': False,
                'strings_to_urls': False,
                'use_zip64': True  # Supporto per file molto grandi
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

    def write_dataframe(self, df: pd.DataFrame, sheet_name: str):
        """Scrive un DataFrame in un foglio Excel in modo ottimizzato"""
        if self._is_closed:
            raise ValueError("Workbook già chiuso")
            
        start_time = time.time()
        total_rows = len(df)
        
        try:
            # Ottimizza il DataFrame
            df = self._optimize_dataframe(df)
            
            # Crea il foglio
            worksheet = self.workbook.add_worksheet(sheet_name)
            
            # Scrivi header
            for col, name in enumerate(df.columns):
                worksheet.write(0, col, str(name), self.formats['header'])
            
            # Calcola e imposta larghezze colonne
            widths = self._calculate_column_widths(df)
            for col, width in widths.items():
                worksheet.set_column(col, col, width)
            
            # Scrivi i dati in chunks
            chunk_size = min(self.CHUNK_SIZE, total_rows)
            for start_idx in range(0, total_rows, chunk_size):
                end_idx = min(start_idx + chunk_size, total_rows)
                chunk = df.iloc[start_idx:end_idx]
                
                # Scrivi ogni riga del chunk
                for i, row in enumerate(chunk.itertuples(index=False), start_idx + 1):
                    for col, value in enumerate(row):
                        if pd.isna(value):
                            worksheet.write(i, col, '', self.formats['base'])
                        elif isinstance(value, pd.Timestamp):
                            worksheet.write_datetime(i, col, value.to_pydatetime(), self.formats['date'])
                        else:
                            worksheet.write(i, col, value, self.formats['base'])
                
                if (start_idx + chunk_size) % (chunk_size * 2) == 0:
                    log_info(f"Scritte {start_idx + chunk_size} righe di {total_rows} per {sheet_name}")
                    gc.collect()
            
            elapsed = time.time() - start_time
            log_info(f"Elaborazione {sheet_name} completata in {elapsed:.2f} secondi")
            
        except Exception as e:
            log_info(f"Errore durante la scrittura di {sheet_name}: {str(e)}")
            raise

    def _optimize_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Ottimizza il DataFrame per l'esportazione
        """
        # Converti le colonne object in categorie dove possibile
        for col in df.select_dtypes(include=['object']).columns:
            if df[col].nunique() / len(df) < 0.5:  # Se meno del 50% dei valori sono unici
                df[col] = df[col].astype('category')
        
        # Ottimizza i tipi di dati numerici
        for col in df.select_dtypes(include=['int64', 'float64']).columns:
            df[col] = pd.to_numeric(df[col], downcast='integer')
        
        return df

    def _calculate_column_widths(self, df: pd.DataFrame, sample_size: int = 1000) -> Dict[int, int]:
        """
        Calcola le larghezze ottimali delle colonne usando un campione di dati
        """
        if df.empty:
            return {}
            
        widths = {}
        # Usa un campione casuale per performance
        sample = df.sample(n=min(sample_size, len(df))) if len(df) > sample_size else df
        
        for idx, col in enumerate(df.columns):
            # Larghezza header
            max_width = len(str(col)) + 2
            
            # Larghezza dati
            col_data = sample.iloc[:, idx].dropna().astype(str)
            if not col_data.empty:
                data_width = col_data.str.len().max() + 2
                max_width = min(max(max_width, data_width), self.MAX_COLUMN_WIDTH)
            
            widths[idx] = max_width
            
        return widths

    def finalize(self):
        """Finalizza il file Excel unendo i file parquet"""
        if self._is_closed:
            log_info("Workbook già chiuso")
            return

        if not self.large_files:
            if self.workbook:
                self.workbook.close()
                self._is_closed = True
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
                    for batch in reader.iter_batches(batch_size=self.CHUNK_SIZE):
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
                        
                        if current_row % (self.CHUNK_SIZE * 2) == 0:
                            log_info(f"Processate {current_row-1} righe per {sheet_name}")
                            gc.collect()
                        
                        del df_chunk
                    
                    # Imposta le larghezze finali delle colonne
                    for col, width in enumerate(max_widths):
                        worksheet.set_column(col, col, min(width, self.MAX_COLUMN_WIDTH))
                    
                    # Rimuovi il file parquet
                    os.remove(parquet_file)
                    log_info(f"File parquet {sheet_name} rimosso")
                    
                except Exception as e:
                    log_info(f"Errore durante l'elaborazione di {sheet_name}: {str(e)}")
                    raise

            # Chiudi il workbook solo se non è già stato chiuso
            if self.workbook and not self._is_closed:
                self.workbook.close()
                self._is_closed = True
            
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
        finally:
            # Assicurati che il workbook sia chiuso in caso di errori
            if self.workbook and not self._is_closed:
                try:
                    self.workbook.close()
                except:
                    pass
                self._is_closed = True
    
    def __del__(self):
        """Cleanup quando l'oggetto viene distrutto"""
        try:
            if hasattr(self, 'executor'):
                self.executor.shutdown(wait=True)
            if self.workbook and not self._is_closed:
                try:
                    self.workbook.close()
                except:
                    pass
                self._is_closed = True
            if hasattr(self, 'temp_dir') and self.temp_dir.exists():
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

def select_output_directory() -> Path:
    """
    Apre una finestra di dialogo per selezionare la directory di output
    
    Returns:
        Path: Percorso selezionato dall'utente o directory corrente se annullato
    """
    root = tk.Tk()
    root.withdraw()  # Nasconde la finestra principale
    root.attributes('-topmost', True)  # Forza la finestra in primo piano
    
    directory = filedialog.askdirectory(
        title='Seleziona la directory per il salvataggio',
        initialdir=str(Path.cwd())
    )
    
    root.destroy()  # Chiude correttamente la finestra Tk
    
    return Path(directory) if directory else Path.cwd()

def format_clickable_path(path: str) -> str:
    return f"\033[94m{path}\033[0m"

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
    # Chiedi all'utente di selezionare la directory di output se non specificata
    if output_path is None:
        base_path = select_output_directory()
        log_info(f"Directory di output selezionata: {format_clickable_path(str(base_path))}")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = str(Path(base_path) / f"export_{environment_name}_{timestamp}.xlsx")

    exporter = None
    try:
        log_info(f"Inizializzazione export per ambiente {environment_name}")
        exporter = Exporter(output_path)

        # Conteggio totale delle viste da elaborare
        total_views = sum(len(db_config["views"]) for db_config in environment_config)
        progress_bar = tqdm(total=total_views, desc="Progresso totale", unit="vista", file=tqdm_logger)

        # Prima elabora tutte le tabelle piccole
        log_info("Fase 1: Elaborazione tabelle piccole")
        for db_config in environment_config:
            fetcher = None
            try:
                fetcher = DatabaseFetcher(db_config["config"])
                if not fetcher.connect():
                    log_info(f"Impossibile connettersi al database {db_config['config'].get('database')}. Salto alla prossima configurazione.")
                    progress_bar.update(len(db_config["views"]))
                    continue

                for view_name in db_config["views"]:
                    if view_name.lower() not in LARGE_TABLES:
                        try:
                            start_time = time.time()
                            df = fetcher.fetch_view_data(view_name)
                            
                            if df is not None and not df.empty:
                                exporter.write_dataframe(df, view_name)
                                elapsed_time = time.time() - start_time
                                log_info(f"Elaborazione {view_name} completata in {elapsed_time:.2f} secondi")
                            else:
                                log_info(f"Nessun dato da elaborare per {view_name}")
                            
                            progress_bar.update(1)
                            progress_bar.set_description(f"Completata {view_name}")
                            
                        except Exception as e:
                            log_info(f"Errore nell'elaborazione di {view_name}: {str(e)}")
                            progress_bar.update(1)
                            
                        finally:
                            # Forza pulizia memoria
                            if 'df' in locals():
                                del df
                            gc.collect()
                            
            finally:
                if fetcher:
                    fetcher.close()

        # Poi elabora le tabelle grandi
        log_info("Fase 2: Elaborazione tabelle grandi")
        for db_config in environment_config:
            fetcher = None
            try:
                fetcher = DatabaseFetcher(db_config["config"])
                if not fetcher.connect():
                    log_info(f"Impossibile connettersi al database {db_config['config'].get('database')}. Salto alla prossima configurazione.")
                    progress_bar.update(len(db_config["views"]))
                    continue

                for view_name in db_config["views"]:
                    if view_name.lower() in LARGE_TABLES:
                        try:
                            start_time = time.time()
                            df = fetcher.fetch_view_data(view_name)
                            
                            if df is not None and not df.empty:
                                exporter.write_dataframe(df, view_name)
                                elapsed_time = time.time() - start_time
                                log_info(f"Elaborazione {view_name} completata in {elapsed_time:.2f} secondi")
                            else:
                                log_info(f"Nessun dato da elaborare per {view_name}")
                            
                            progress_bar.update(1)
                            progress_bar.set_description(f"Completata {view_name}")
                            
                        except Exception as e:
                            log_info(f"Errore nell'elaborazione di {view_name}: {str(e)}")
                            progress_bar.update(1)
                            
                        finally:
                            # Forza pulizia memoria
                            if 'df' in locals():
                                del df
                            gc.collect()
                            
            finally:
                if fetcher:
                    fetcher.close()

        progress_bar.close()
        
        # Fase finale: finalizzazione del file
        log_info("Fase 3: Finalizzazione del file Excel")
        exporter.finalize()
        log_info(f"Export completato. File salvato in: {format_clickable_path(output_path)}")
        return output_path
        
    except Exception as e:
        log_info(f"Errore durante l'export: {str(e)}")
        raise
    finally:
        # Assicurati che l'exporter venga chiuso correttamente
        if exporter and not exporter._is_closed:
            try:
                exporter.finalize()
            except:
                pass
        gc.collect()
