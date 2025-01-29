from datetime import datetime
from pathlib import Path
import pandas as pd
from typing import Dict, List
import xlsxwriter
import logging
from src.excel.reader import ExcelReader
from src.excel.exporter import DatabaseFetcher, select_output_directory, format_clickable_path
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
    
    def export_errors_to_excel(self, output_dir: str = None, environment_name: str = "UNKNOWN") -> str:
        """
        Esporta gli errori in un file Excel con formattazione migliorata
        
        Args:
            output_dir: Directory di output per il file Excel. Se None, verrà chiesto all'utente
            environment_name: Nome dell'ambiente (es. SIT, UAT, PREPROD)
        
        Returns:
            str: Percorso del file Excel creato
        """
        if not self.get_all_errors():
            return ""
            
        # Se non è specificata una directory di output, chiedi all'utente
        if output_dir is None:
            output_dir = select_output_directory()
            
        # Crea il nome del file con timestamp e ambiente
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = Path(output_dir) / f"dq_errors_{environment_name}_{timestamp}.xlsx"
        
        # Crea il workbook con xlsxwriter per un maggior controllo sulla formattazione
        workbook = None
        try:
            workbook = pd.ExcelWriter(output_file, engine='xlsxwriter')
            wb = workbook.book
            
            # Definisci gli stili
            header_format = wb.add_format({
                'bold': True,
                'font_size': 12,
                'bg_color': '#4F81BD',
                'font_color': 'white',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'text_wrap': True
            })
            
            error_format = wb.add_format({
                'bg_color': '#FFC7CE',
                'font_color': '#9C0006'
            })
            
            normal_format = wb.add_format({
                'font_size': 11,
                'align': 'left',
                'valign': 'vcenter',
                'text_wrap': True
            })
            
            # Crea fogli dettagliati per ogni regola
            for rule_name, rule_errors in self.get_all_errors().items():
                if not rule_errors:
                    continue
                    
                # Crea DataFrame con gli errori
                df = pd.DataFrame(rule_errors)
                
                # Rinomina le colonne di base
                column_renames = {
                    'sheet_name': 'Foglio',
                    'row_index': 'Riga Excel'
                }
                
                # Rinomina le colonne mantenendo le colonne dei campi chiave invariate
                df = df.rename(columns=column_renames)
                
                # Riordina le colonne: prima Foglio e Riga Excel, poi i campi chiave
                fixed_columns = ['Foglio', 'Riga Excel']
                key_columns = [col for col in df.columns if col not in fixed_columns]
                df = df[fixed_columns + key_columns]
                
                # Ordina per foglio e riga
                df = df.sort_values(['Foglio', 'Riga Excel'])
                
                # Scrivi il DataFrame
                sheet_name = rule_name[:31]  # Excel limita i nomi dei fogli a 31 caratteri
                df.to_excel(workbook, sheet_name=sheet_name, index=False)
                
                # Ottieni il worksheet
                worksheet = wb.get_worksheet_by_name(sheet_name)
                
                # Formatta l'header
                for col_num, col in enumerate(df.columns):
                    worksheet.write(0, col_num, col, header_format)
                    
                    # Imposta larghezza colonna basata sul contenuto
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(col)
                    )
                    worksheet.set_column(col_num, col_num, min(max_length + 2, 50))
                
                # Formatta le celle
                for row_num in range(1, len(df) + 1):
                    for col_num, value in enumerate(df.iloc[row_num-1]):
                        if pd.isna(value):
                            worksheet.write(row_num, col_num, '', normal_format)
                        else:
                            worksheet.write(row_num, col_num, value, 
                                         error_format if col_num >= len(fixed_columns) else normal_format)
                
                # Aggiungi filtri
                worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
                
                # Congela le prime due colonne e la prima riga
                worksheet.freeze_panes(1, 2)
            
            # Salva il file
            workbook.close()
            
            return str(output_file)
            
        except Exception as e:
            # In caso di errore, assicurati che il workbook sia chiuso
            if workbook:
                try:
                    workbook.close()
                except:
                    pass
            raise e
