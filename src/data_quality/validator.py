from typing import Dict, List, Type
import pandas as pd
from pathlib import Path
import logging
from datetime import datetime

from src.excel.reader import ExcelReader
from src.excel.exporter import DatabaseFetcher
from .base import DQRule
from .rules.duplicate_check import DuplicateCheckRule

class DataQualityValidator:
    """Gestore principale per la validazione della qualità dei dati"""
    
    def __init__(self):
        self.rules: List[DQRule] = []
        self.excel_data: Dict[str, pd.DataFrame] = {}
        self.db_data: Dict[str, pd.DataFrame] = {}
        
    def add_rule(self, rule: DQRule):
        """Aggiunge una regola di validazione"""
        self.rules.append(rule)
        
    def load_excel_data(self, excel_path: str):
        """Carica i dati dal file Excel"""
        reader = ExcelReader(excel_path)
        reader.validate_sheets()
        self.excel_data = reader.read_all_sheets()
        
    def load_db_data(self, env_config: List[Dict]):
        """Carica i dati dal database"""
        self.db_data = {}
        
        for config in env_config:
            db_config = config['config']
            views = config['views']
            
            # Crea un'istanza di DatabaseFetcher per questa configurazione
            fetcher = DatabaseFetcher(config=db_config)
            
            try:
                # Connettiti al database
                fetcher.connect()
                
                # Recupera i dati per ogni vista
                for view in views:
                    df = fetcher.fetch_view_data(view)
                    if df is not None:
                        self.db_data[view] = df
                        
            finally:
                # Chiudi sempre la connessione
                fetcher.close()
    
    def validate(self) -> bool:
        """Esegue tutte le regole di validazione"""
        if not self.excel_data or not self.db_data:
            raise ValueError("Dati Excel o DB non caricati")
            
        all_valid = True
        for rule in self.rules:
            if not rule.validate(self.excel_data, self.db_data):
                all_valid = False
                
        return all_valid
    
    def get_all_errors(self) -> Dict[str, List[Dict]]:
        """Raccoglie tutti gli errori da tutte le regole"""
        errors = {}
        for rule in self.rules:
            if rule.errors:
                errors[rule.name] = rule.get_errors()
        return errors
    
    def export_errors_to_excel(self, output_dir: str = "DQ_Results") -> str:
        """Esporta gli errori in un file Excel"""
        if not self.get_all_errors():
            return ""
            
        # Crea directory se non esiste
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        # Crea il nome del file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = output_path / f"dq_errors_{timestamp}.xlsx"
        
        # Prepara i dati per l'export
        with pd.ExcelWriter(output_file) as writer:
            for rule_name, rule_errors in self.get_all_errors().items():
                if rule_errors:
                    df = pd.DataFrame(rule_errors)
                    df.to_excel(writer, sheet_name=rule_name, index=False)
        
        return str(output_file)
