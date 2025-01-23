import os
from typing import Dict, Optional
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm
import logging

from config.constants import REQUIRED_SHEETS, EXCEL_ENGINE, DEFAULT_DTYPE
from src.utils.helpers import clean_dataframe

class ExcelReader:
    def __init__(self, file_path: str):
        """Initialize ExcelReader with file path."""
        self.file_path = file_path
        self.excel_file = pd.ExcelFile(file_path, engine=EXCEL_ENGINE)
        
    def validate_sheets(self) -> None:
        """Validate that all required sheets exist."""
        missing_sheets = set(REQUIRED_SHEETS) - set(self.excel_file.sheet_names)
        if missing_sheets:
            raise ValueError(f"Fogli mancanti: {missing_sheets}")
    
    def read_single_sheet(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """Read a single sheet from the Excel file with optimizations."""
        try:
            df = pd.read_excel(
                self.excel_file,
                sheet_name=sheet_name,
                dtype=DEFAULT_DTYPE
            )
            
            df = clean_dataframe(df)
            
            if df.empty:
                logging.warning(f"Nessun dato trovato nel foglio: {sheet_name}")
                return None
                
            return df
            
        except Exception as e:
            logging.error(f"Errore nella lettura del foglio {sheet_name}: {str(e)}")
            return None
    
    def read_all_sheets(self) -> Dict[str, pd.DataFrame]:
        """Read all required sheets in parallel."""
        sheets_data = {}
        max_workers = min(len(REQUIRED_SHEETS), os.cpu_count() * 2)
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(self.read_single_sheet, sheet): sheet 
                for sheet in REQUIRED_SHEETS
            }
            
            with tqdm(total=len(REQUIRED_SHEETS), desc="Lettura fogli") as pbar:
                for future in futures:
                    sheet_name = futures[future]
                    try:
                        df = future.result()
                        if df is not None and not df.empty:
                            sheets_data[sheet_name] = df
                        pbar.update(1)
                    except Exception as e:
                        logging.error(f"Errore nell'elaborazione del foglio {sheet_name}: {str(e)}")
                        pbar.update(1)
        
        return sheets_data
    
    def read_excel_file(self) -> Dict[str, pd.DataFrame]:
        """Main method to read Excel file with validation."""
        self.validate_sheets()
        return self.read_all_sheets()
