from typing import Dict, List
import pandas as pd
import time
import logging
import re
import sys
from tqdm import tqdm
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
            
            # Filtra il DB in base agli ExerciseType presenti nelle righe da inserire
            if 'Exercise Type' in insert_rows.columns and 'Exercise Type' in db_df.columns:
                unique_exercise_types = insert_rows['ExerciseType'].dropna().unique()
                if len(unique_exercise_types) > 0:
                    logging.info(f"Filtraggio DB per ExerciseType: trovati {len(unique_exercise_types)} tipi unici")
                    print(f"Filtraggio DB per ExerciseType: {len(unique_exercise_types)} tipi unici")
                    
                    # Converti i valori a stringa per garantire la compatibilità
                    unique_exercise_types_str = [str(et).strip() for et in unique_exercise_types if not pd.isna(et)]
                    
                    # Filtra il DataFrame del DB per includere solo le righe con gli ExerciseType corrispondenti
                    original_db_size = len(db_df)
                    db_df = db_df[db_df['Exercise Type'].astype(str).str.strip().isin(unique_exercise_types_str)]
                    filtered_db_size = len(db_df)
                    
                    reduction_percentage = ((original_db_size - filtered_db_size) / original_db_size * 100) if original_db_size > 0 else 0
                    logging.info(f"DB filtrato: da {original_db_size} a {filtered_db_size} righe (-{reduction_percentage:.2f}%)")
                    print(f"DB filtrato: da {original_db_size} a {filtered_db_size} righe (-{reduction_percentage:.2f}%)")
            
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
        
        print(f"\nValidazione duplicati per {sheet_name} - {total_rows} righe da controllare")
        
        # Pre-converti le colonne del DB per evitare conversioni ripetute
        print("Pre-elaborazione dati del database in corso...")
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
        
        # Inizializza il contatore per i duplicati trovati
        duplicates_found = 0
        
        # Processa le righe Excel in chunk con barra di progresso
        chunks = [(i, min(i + self.chunk_size, total_rows)) for i in range(0, total_rows, self.chunk_size)]
        
        # Crea una barra di progresso per i chunk
        with tqdm(total=len(chunks), desc="Elaborazione chunk", unit="chunk", bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} chunk [{elapsed}<{remaining}, {rate_fmt}]") as chunk_pbar:
            for chunk_idx, (chunk_start, chunk_end) in enumerate(chunks):
                chunk = insert_rows.iloc[chunk_start:chunk_end]
                chunk_size = len(chunk)
                
                # Aggiorna la descrizione della barra di progresso con informazioni sul chunk corrente
                chunk_pbar.set_description(f"Chunk {chunk_idx+1}/{len(chunks)} [{chunk_start}-{chunk_end}]")
                
                # Crea una barra di progresso per le righe all'interno del chunk
                with tqdm(total=chunk_size, desc="Righe", unit="riga", leave=False, 
                         bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} righe [{elapsed}<{remaining}, {rate_fmt}]") as row_pbar:
                    
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
                            duplicates_found += 1
                            
                            # Aggiorna la descrizione della barra di progresso con il numero di duplicati trovati
                            row_pbar.set_postfix(duplicati=duplicates_found)
                            
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
                        
                        # Aggiorna la barra di progresso delle righe
                        row_pbar.update(1)
                        
                        # Libera memoria ogni 100 righe
                        if idx % 100 == 0:
                            import gc
                            gc.collect()
                
                # Aggiorna la barra di progresso dei chunk
                chunk_pbar.update(1)
                
                # Stampa un riepilogo del chunk completato
                print(f"Chunk {chunk_idx+1}/{len(chunks)} completato - Trovati {duplicates_found} duplicati finora")
        
        # Stampa un riepilogo finale
        print(f"\nValidazione completata per {sheet_name}:")
        print(f"- Righe controllate: {total_rows}")
        print(f"- Duplicati trovati: {duplicates_found}")
        
        return has_errors
