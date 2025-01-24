import logging
from datetime import datetime
from pathlib import Path
from src.excel.reader import ExcelReader
from src.query.generator import QueryGenerator
from src.excel.exporter import export_to_excel
from src.utils.helpers import setup_logging, validate_file_path, extract_id_from_path, select_excel_file
from config.constants import ENVIRONMENTS, OUTPUT_DIR

def generate_queries():
    """Funzione per generare le query SQL da un file Excel"""
    try:
        # Get the Excel file path using file dialog
        try:
            excel_path = select_excel_file()
            print(f"\nFile selezionato: {excel_path}")
        except ValueError as e:
            print(f"\nErrore: {str(e)}")
            return False

        # Extract ID from filename if present
        file_id = extract_id_from_path(str(excel_path))

        # Read Excel data
        print("\nLettura del file Excel in corso...")
        excel_reader = ExcelReader(str(excel_path))
        excel_reader.validate_sheets()  # Validate required sheets exist
        data = excel_reader.read_all_sheets()  # Read all sheets in parallel

        if not data:
            raise ValueError("Nessun dato valido trovato nel file Excel")

        # Generate queries
        print("\nGenerazione delle query in corso...")
        query_generator = QueryGenerator(data, str(excel_path), file_id)
        query_generator.generate_all_queries()
        output_file = query_generator.save_queries()

        print(f"\nQuery salvate nel file: {output_file}")
        logging.info(f"Query generate con successo nel file: {output_file}")
        return True
        
    except Exception as e:
        logging.error(f"Si è verificato un errore durante la generazione delle query: {str(e)}")
        return False

def extract_data():
    """Funzione per estrarre i dati dal database"""
    try:
        while True:
            print(f"\nAmbienti disponibili: {', '.join(ENVIRONMENTS.keys())}")
            env_choice = input("Seleziona ambiente: ").upper()
            if env_choice in ENVIRONMENTS:
                break
            print(f"\nAmbiente non valido. Scegli tra: {', '.join(ENVIRONMENTS.keys())}")

        # Export database state to Excel
        print("\nEstrazione dati in corso...")
        export_path = export_to_excel(ENVIRONMENTS[env_choice], environment_name=env_choice)
        print(f"\nDati esportati nel file: {export_path}")
        return True

    except Exception as e:
        logging.error(f"Si è verificato un errore durante l'estrazione dei dati: {str(e)}")
        return False

def show_menu():
    """Mostra il menu principale e restituisce la scelta dell'utente"""
    while True:
        print("\n=== Prometheus Query Generator ===")
        print("=================================")
        print("\nOperazioni disponibili:")
        print("1. Genera query SQL da file Excel")
        print("2. Estrai dati dal database")
        print("3. Esci")
        
        choice = input("\nScegli operazione (1-3): ")
        if choice in ['1', '2', '3']:
            return choice
        print("\n✗ Scelta non valida. Seleziona un numero tra 1 e 4.")

def main():
    setup_logging()
    
    while True:
        choice = show_menu()
        
        if choice == '1':  # Solo generazione query
            if generate_queries():
                print("\n✓ Generazione query completata con successo!")
            else:
                print("\n✗ Si è verificato un errore durante la generazione delle query.")
                
        elif choice == '2':  # Solo estrazione dati
            if extract_data():
                print("\n✓ Estrazione dati completata con successo!")
            else:
                print("\n✗ Si è verificato un errore durante l'estrazione dei dati.")
        else:  # Esci
            print("\nGrazie per aver utilizzato Prometheus Query Generator!")
            break

if __name__ == "__main__":
    main()
