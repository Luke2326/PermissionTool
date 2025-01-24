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
    """Extract ID from file path if present."""
    match = re.search(r"ID_\d+", file_path)
    if match:
        id_value = match.group()
        logging.info(f"ID trovato: {id_value}")
        return id_value
    logging.info("Nessun ID trovato.")
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
    """Generate a timestamp string for file naming."""
    return datetime.now().strftime("%Y%m%d_%H%M%S")

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
