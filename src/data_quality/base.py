from abc import ABC, abstractmethod
from typing import Dict, List
import pandas as pd

class DQRule(ABC):
    """Interfaccia base per le regole di Data Quality"""
    
    def __init__(self, name: str, description: str):
        self.name = name
        self.description = description
        self.errors: List[Dict] = []
    
    @abstractmethod
    def validate(self, excel_data: Dict[str, pd.DataFrame], db_data: Dict[str, pd.DataFrame]) -> bool:
        """
        Esegue la validazione dei dati.
        
        Args:
            excel_data: Dict con i dati del file Excel {sheet_name: DataFrame}
            db_data: Dict con i dati del database {view_name: DataFrame}
            
        Returns:
            bool: True se la validazione passa, False altrimenti
        """
        pass
    
    def add_error(self, sheet_name: str, row_index: int, message: str):
        """Aggiunge un errore alla lista degli errori"""
        self.errors.append({
            'sheet_name': sheet_name,
            'row_index': row_index,
            'message': message
        })
    
    def get_errors(self) -> List[Dict]:
        """Restituisce la lista degli errori"""
        return self.errors
    
    def clear_errors(self):
        """Pulisce la lista degli errori"""
        self.errors = []
