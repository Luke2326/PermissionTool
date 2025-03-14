from typing import Dict, List
import pandas as pd
import time
import logging
import re
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
        # Soglia per decidere se usare l'approccio a chunk
        self.large_dataset_threshold = 10000
        # Dimensione del chunk per dataset grandi
        self.chunk_size = 5000
    
    def validate(self, excel_data: Dict[str, pd.DataFrame], db_data: Dict[str, pd.DataFrame]) -> bool:
        """
        Verifica la presenza di duplicati tra i dati Excel e quelli del DB.
        Controlla solo le righe con Delta = 'INSERT'.
        
        Returns:
            bool: True se non ci sono errori, False altrimenti
        """
        start_time = time.time()
        self.clear_errors()
        has_errors = False
        
        for sheet_name, excel_df in excel_data.items():
            if sheet_name not in self.key_mapping:
                continue
            
            sheet_start_time = time.time()
            logging.info(f"Validazione duplicati per il foglio: {sheet_name}")
                
            # Filtra solo le righe con Delta = INSERT
            insert_rows = excel_df[excel_df['Delta'].str.upper() == 'INSERT']
            if insert_rows.empty:
                logging.info(f"Nessuna riga da inserire nel foglio {sheet_name}")
                continue
                
            db_df = db_data.get(sheet_name)
            if db_df is None:
                logging.info(f"Nessun dato DB trovato per il foglio {sheet_name}")
                continue
                
            key_columns = self.key_mapping[sheet_name]
            logging.info(f"Colonne chiave per {sheet_name}: {key_columns}")
            
            # Verifica se tutte le colonne chiave esistono
            missing_keys_excel = [k for k in key_columns if k not in insert_rows.columns]
            missing_keys_db = [k for k in key_columns if k not in db_df.columns]
            
            if missing_keys_excel or missing_keys_db:
                logging.warning(f"Colonne chiave mancanti - Excel: {missing_keys_excel}, DB: {missing_keys_db}")
                continue
            
            # Scegli l'approccio in base alla dimensione dei dataset
            if len(db_df) > self.large_dataset_threshold or len(insert_rows) > 1000:
                logging.info(f"Utilizzando approccio a chunk per {sheet_name} (righe DB: {len(db_df)}, righe Excel: {len(insert_rows)})")
                has_errors |= self._validate_large_dataset(insert_rows, db_df, key_columns, sheet_name)
            else:
                logging.info(f"Utilizzando approccio vettoriale per {sheet_name} (righe DB: {len(db_df)}, righe Excel: {len(insert_rows)})")
                has_errors |= self._validate_small_dataset(insert_rows, db_df, key_columns, sheet_name)
            
            sheet_end_time = time.time()
            logging.info(f"Validazione {sheet_name} completata in {sheet_end_time - sheet_start_time:.2f} secondi")
        
        end_time = time.time()
        logging.info(f"Validazione duplicati completata in {end_time - start_time:.2f} secondi")
        return not has_errors
    
    def _validate_small_dataset(self, insert_rows: pd.DataFrame, db_df: pd.DataFrame, 
                               key_columns: List[str], sheet_name: str) -> bool:
        """
        Metodo ottimizzato per dataset piccoli usando operazioni vettoriali.
        
        Returns:
            bool: True se sono stati trovati errori, False altrimenti
        """
        has_errors = False
        
        # Pre-converti le colonne del DB per evitare conversioni ripetute
        db_df_converted = {}
        for key in key_columns:
            if key in db_df.columns:
                # Converti solo se necessario
                db_df_converted[key] = db_df[key].astype(str).str.strip() if db_df[key].dtype != 'object' else db_df[key].str.strip()
        
        # Verifica ogni riga da inserire
        for idx, row in insert_rows.iterrows():
            matches = pd.Series([True] * len(db_df), index=db_df.index)
            
            for key in key_columns:
                if key not in row or key not in db_df:
                    continue
                    
                excel_value = row[key]
                
                # Gestione valori nulli
                if pd.isna(excel_value):
                    matches &= pd.isna(db_df[key])
                else:
                    # Usa i valori pre-convertiti
                    excel_str = str(excel_value).strip()
                    matches &= (db_df_converted[key] == excel_str)
            
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
        
        return has_errors
    
    def _validate_large_dataset(self, insert_rows: pd.DataFrame, db_df: pd.DataFrame, 
                               key_columns: List[str], sheet_name: str) -> bool:
        """
        Metodo ottimizzato per dataset grandi usando elaborazione a chunk.
        
        Returns:
            bool: True se sono stati trovati errori, False altrimenti
        """
        has_errors = False
        total_rows = len(insert_rows)
        
        # Pre-converti le colonne del DB per evitare conversioni ripetute
        db_df_converted = {}
        for key in key_columns:
            if key in db_df.columns:
                try:
                    # Converti solo se necessario
                    if db_df[key].dtype != 'object':
                        db_df_converted[key] = db_df[key].astype(str).str.strip()
                    else:
                        db_df_converted[key] = db_df[key].str.strip()
                except Exception as e:
                    logging.warning(f"Errore nella pre-conversione della colonna {key}: {str(e)}")
                    # Fallback: non pre-convertire questa colonna
        
        # Processa le righe Excel in chunk
        for chunk_start in range(0, total_rows, self.chunk_size):
            chunk_end = min(chunk_start + self.chunk_size, total_rows)
            chunk = insert_rows.iloc[chunk_start:chunk_end]
            
            logging.info(f"Processando chunk {chunk_start}-{chunk_end} di {total_rows} righe")
            
            # Processa ogni riga nel chunk
            for idx, row in chunk.iterrows():
                # Inizializza una maschera booleana per tutte le righe del DB
                matches = pd.Series([True] * len(db_df), index=db_df.index)
                
                # Applica il filtro per ogni colonna chiave
                for key in key_columns:
                    if key not in row or key not in db_df.columns:
                        continue
                    
                    excel_value = row[key]
                    
                    # Gestione valori nulli
                    if pd.isna(excel_value):
                        matches &= pd.isna(db_df[key])
                    else:
                        excel_str = str(excel_value).strip()
                        
                        # Usa i valori pre-convertiti se disponibili
                        if key in db_df_converted:
                            matches &= (db_df_converted[key] == excel_str)
                        else:
                            # Fallback al metodo tradizionale
                            db_values = db_df[key]
                            db_str = db_values.astype(str).str.strip()
                            matches &= (db_str == excel_str)
                
                # Verifica se ci sono corrispondenze
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
                
                # Libera memoria ogni 100 righe
                if idx % 100 == 0:
                    import gc
                    gc.collect()
        
        return has_errors
