from typing import Dict, List
import pandas as pd
from ..base import DQRule

class DuplicateCheckRule(DQRule):
    """Regola che verifica la presenza di duplicati tra Excel e DB"""
    
    def __init__(self):
        super().__init__(
            name="DuplicateCheck",
            description="Verifica la presenza di record già esistenti nel database"
        )
        # Mapping delle chiavi primarie per ogni sheet/vista
        self.key_mapping = {
            'Prometheus Entities': ['Entity', 'ObjectType','Entity Type','Country Code','Node Level'],
            'Prometheus Data Items': ['Entity','File Type'],
            'Prometheus Groups': ['Exercise Type','Group Code','Reference Node','Description'],
            'Prometheus Permissions': ['Exercise Type','Group Code','Functionality Name','Entity Name','FileType Name'],
            'Prometheus Roles': ['Role Unique Name','Role Name'],
            'Prometheus Permission Set': ['ProfileSetName', 'ProfileSetVersionName', 'Version Number', 'Set Version Name'],
            'Prometheus Set Role Group Ver': ['Exercise Type','Set Version Name','Role Unique Name','Group Unique Name']
        }
    
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
                match_condition = True
                for key in key_columns:
                    if key in row and key in db_df:
                        match_condition &= (db_df[key] == row[key])
                
                if match_condition.any():
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
