from typing import Dict, List
import pandas as pd
from ..base import DQRule
from config.constants import KEY_MAPPING

class DuplicateCheckRule(DQRule):
    """Regola che verifica la presenza di duplicati tra Excel e DB"""
    
    def __init__(self):
        super().__init__(
            name="DuplicateCheck",
            description="Verifica la presenza di record già esistenti nel database"
        )
        self.key_mapping = KEY_MAPPING
    
    def validate(self, excel_data: Dict[str, pd.DataFrame], db_data: Dict[str, pd.DataFrame]) -> bool:
        """
        Verifica la presenza di duplicati tra i dati Excel e quelli del DB.
        Controlla solo le righe con Delta = 'INSERT'.
        """
        self.clear_errors()
        has_errors = False
        
        for sheet_name, excel_df in excel_data.items():
            if sheet_name not in self.key_mapping:
                continue
                
            # Filtra solo le righe con Delta = INSERT
            insert_rows = excel_df[excel_df['Delta'].str.upper() == 'INSERT']
            if insert_rows.empty:
                continue
                
            db_df = db_data.get(sheet_name)
            if db_df is None:
                continue
                
            key_columns = self.key_mapping[sheet_name]
            
            # Verifica ogni riga da inserire
            for idx, row in insert_rows.iterrows():
                matches = pd.Series([True] * len(db_df), index=db_df.index)
                
                for key in key_columns:
                    if key not in row or key not in db_df:
                        continue
                        
                    excel_value = row[key]
                    db_values = db_df[key]
                    
                    # Gestione valori nulli
                    if pd.isna(excel_value):
                        matches &= pd.isna(db_values)
                    else:
                        # Converti a stringa per gestire tipi diversi
                        excel_str = str(excel_value).strip()
                        db_str = db_values.astype(str).str.strip()
                        matches &= (db_str == excel_str)
                
                if matches.any():
                    has_errors = True
                    # Crea un dizionario con i valori delle chiavi
                    error_details = {
                        'sheet_name': sheet_name,
                        'row_index': idx + 2,  # +2 per compensare l'header e l'indice 0-based
                    }
                    
                    # Aggiungi i valori delle chiavi come campi separati
                    for key in key_columns:
                        if key in row:
                            error_details[key] = row[key]
                    
                    self.errors.append(error_details)
        
        return not has_errors
