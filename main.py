import logging
from datetime import datetime
from pathlib import Path
from src.excel.reader import ExcelReader
from src.query.generator import QueryGenerator
from src.excel.exporter import export_to_excel
from src.utils.helpers import setup_logging, validate_file_path, extract_id_from_path
from config.constants import ENVIRONMENTS, OUTPUT_DIR

def main():
    try:
        # Setup logging
        setup_logging()
        
        # Get the input file path from user
        file_path = input("Inserisci il percorso del file Excel: ").strip()
        
        # Validate input file
        validate_file_path(file_path)
        
        # Extract ID from filename if present
        id = extract_id_from_path(file_path)
        
        # Step 1: Read the Excel file
        logging.info("Lettura del file Excel...")
        excel_reader = ExcelReader(file_path)
        sheets_data = excel_reader.read_excel_file()
        
        if not sheets_data:
            raise ValueError("Nessun dato valido trovato nel file Excel")
        
        # Step 2: Generate queries
        logging.info("Generazione delle query...")
        query_gen = QueryGenerator(sheets_data, file_path, id)
        query_gen.generate_all_queries()
        output_file = query_gen.save_queries()
        
        logging.info(f"Query generate con successo nel file: {output_file}")
        
        # Ask user if they want to export current database state
        while True:
            export_choice = input("\nVuoi esportare lo stato attuale del database? (si/no): ").lower()
            if export_choice in ['si', 'no']:
                break
            print("Per favore, inserisci 'si' o 'no'")

        if export_choice == 'si':
            while True:
                env_choice = input(f"\nSeleziona ambiente ({', '.join(ENVIRONMENTS.keys())}): ").upper()
                if env_choice in ENVIRONMENTS:
                    break
                print(f"Per favore, seleziona un ambiente valido: {', '.join(ENVIRONMENTS.keys())}")

            # Export database state to Excel
            export_path = export_to_excel(ENVIRONMENTS[env_choice])
            print(f"\nStato del database esportato in: {export_path}")

    except Exception as e:
        logging.error(f"Si è verificato un errore: {str(e)}")
        raise

if __name__ == "__main__":
    main()
