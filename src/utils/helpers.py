import os
import logging
from typing import Optional
import re
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from pathlib import Path

def setup_logging() -> None:
    """Configure logging for the application."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

def validate_file_path(file_path: str) -> None:
    """Validate if the file exists and has correct extension."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File non trovato: {file_path}")
    
    if not file_path.endswith('.xlsb'):
        raise ValueError("Il file deve essere in formato .xlsb")
    
    file_size = os.path.getsize(file_path)
    if file_size == 0:
        raise ValueError("Il file è vuoto")
    
    file_size_mb = file_size / (1024 * 1024)
    logging.info(f"Dimensione del file: {file_size_mb:.2f} MB")

def extract_id_from_path(file_path: str) -> Optional[str]:
    """
    Extract ID from file path if present.
    
    Supporta diversi formati:
    - [12345] nel nome del file
    - ID_12345 nel percorso
    - Qualsiasi numero tra parentesi quadre
    
    Args:
        file_path: Percorso del file da cui estrarre l'ID
        
    Returns:
        Optional[str]: ID estratto o None se non trovato
    """
    # Cerca pattern [12345] nel nome del file
    base_name = os.path.basename(file_path)
    match = re.search(r"\[(\d+)\]", base_name)
    if match:
        id_value = match.group(1)
        logging.info(f"ID trovato nel nome del file: {id_value}")
        return id_value
    
    # Cerca pattern ID_12345 nel percorso (compatibilità con versioni precedenti)
    match = re.search(r"ID_(\d+)", file_path)
    if match:
        id_value = match.group(1)
        logging.info(f"ID trovato nel percorso: {id_value}")
        return id_value
    
    logging.info("Nessun ID trovato nel file.")
    return None

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Clean and prepare DataFrame for processing."""
    # Filter rows where Delta is not empty/null
    df = df[df['Delta'].notna()].copy()
    
    # Clean string columns efficiently using vectorized operations
    object_columns = df.select_dtypes(include=['object']).columns
    df[object_columns] = df[object_columns].apply(lambda x: x.str.strip())
    
    return df

def generate_timestamp() -> str:
    """Generate a timestamp string for file naming in format YYYY_MM_DD."""
    return datetime.now().strftime("%Y_%m_%d")

def extract_excel_filename(file_path: str) -> str:
    """
    Extract the filename without extension from a file path.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        str: Filename without extension
    """
    base_name = os.path.basename(file_path)
    return os.path.splitext(base_name)[0]

def get_next_version(output_dir: str, base_filename: str, date_str: str) -> str:
    """
    Determine the next version number for a file.
    
    Args:
        output_dir: Directory where files are stored
        base_filename: Base name of the file without date and version
        date_str: Date string in format YYYY_MM_DD
        
    Returns:
        str: Version string in format 'vXX' (e.g., 'v01', 'v02')
    """
    # Se la directory non esiste, restituisci v01
    if not os.path.exists(output_dir):
        return "v01"
        
    # Pattern to match files with the same base name and date
    pattern = f"{base_filename} {date_str} v(\\d+)\\.sql$"
    
    # Find all matching files
    max_version = 0
    try:
        for filename in os.listdir(output_dir):
            match = re.search(pattern, filename)
            if match:
                version = int(match.group(1))
                max_version = max(max_version, version)
    except (FileNotFoundError, PermissionError) as e:
        logging.warning(f"Errore durante la lettura della directory {output_dir}: {str(e)}")
        return "v01"
    
    # Return next version
    return f"v{(max_version + 1):02d}"

def ensure_output_directory(base_dir: str) -> str:
    """Ensure output directory exists and return its path."""
    output_dir = os.path.join(os.path.expanduser("~"), "Documents", base_dir)
    os.makedirs(output_dir, exist_ok=True)
    return output_dir

def select_excel_file() -> Path:
    """
    Apre una finestra di dialogo per selezionare un file Excel
    
    Returns:
        Path: Percorso del file Excel selezionato
        
    Raises:
        ValueError: Se nessun file è stato selezionato
    """
    root = tk.Tk()
    root.withdraw()  # Nasconde la finestra principale
    root.attributes('-topmost', True)  # Forza la finestra in primo piano
    
    file_path = filedialog.askopenfilename(
        title='Seleziona il file Excel',
        initialdir=str(Path.cwd()),
        filetypes=[
            ('Excel files', '*.xlsx'),
            ('Excel files', '*.xlsb'),
            ('Excel files', '*.xls'),
            ('All files', '*.*')
        ]
    )
    
    root.destroy()  # Chiude correttamente la finestra Tk
    
    if not file_path:
        raise ValueError("Nessun file selezionato")
        
    return Path(file_path)

def format_clickable_path(path: str) -> str:
    """
    Formatta un percorso file come link cliccabile nel terminale con colori
    
    Args:
        path: Il percorso del file da formattare
        
    Returns:
        str: Il percorso formattato come link cliccabile e colorato
    """
    # Codici ANSI per i colori e lo stile
    BLUE = '\033[94m'
    UNDERLINE = '\033[4m'
    END = '\033[0m'
    
    # Crea il link cliccabile (OSC 8) con colore blu e sottolineato
    clickable = f"\033]8;;file://{path}\033\\{BLUE}{UNDERLINE}{path}{END}\033]8;;\033\\"
    
    return clickable
