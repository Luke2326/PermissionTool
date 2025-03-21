"""
Test script per verificare la generazione dei file SQL per ogni ambiente.
"""
import os
import pandas as pd
from datetime import datetime
from src.query.generator import QueryGenerator
from config.constants import ENVIRONMENTS, OUTPUT_DIR
from src.utils.helpers import ensure_output_directory, generate_timestamp

def test_save_queries():
    """Test per verificare la generazione dei file SQL per ogni ambiente."""
    print("Inizio test per la generazione dei file SQL...")
    
    # Crea un DataFrame di test
    test_data = {
        'Prometheus Entities': pd.DataFrame({
            'Entity': ['Test Entity'],
            'ObjectType': ['Test Type'],
            'Entity Type': ['Test Entity Type'],
            'Country Code': ['IT'],
            'Node Level': [1],
            'Delta': ['I']
        })
    }
    
    # Crea un'istanza di QueryGenerator con i dati di test
    file_path = "test_file.xlsx"
    file_id = "12345"
    query_generator = QueryGenerator(test_data, file_path, file_id)
    
    # Aggiungi una query di test
    query_generator.queries = ["INSERT INTO test_table (id, name) VALUES (1, 'Test')"]
    
    # Nome della persona che esegue lo script (per test)
    executed_by = "Test User"
    
    # Genera i file SQL
    output_files = query_generator.save_queries(executed_by)
    
    # Visualizza i percorsi completi
    print("\nPercorsi completi dei file generati:")
    for output_file in output_files:
        print(f"- {output_file}")
    
    # Verifica che siano stati generati 4 file (uno per ogni ambiente)
    if len(output_files) == 4:
        print(f"✅ Generati correttamente {len(output_files)} file SQL (uno per ogni ambiente)")
    else:
        print(f"❌ Errore: generati {len(output_files)} file invece di 4")
    
    # Verifica che ogni file contenga il nome dell'ambiente
    for env_name in ENVIRONMENTS.keys():
        found = False
        for output_file in output_files:
            if f"[{env_name}]" in output_file:
                found = True
                print(f"✅ File per l'ambiente {env_name} generato correttamente: {os.path.basename(output_file)}")
                
                # Verifica il contenuto del file
                with open(output_file, 'r') as f:
                    content = f.read()
                    
                    # Verifica che il file contenga il controllo dell'IP
                    if "inet_server_addr()" in content:
                        print(f"  ✅ Il file contiene il controllo dell'IP")
                    else:
                        print(f"  ❌ Il file non contiene il controllo dell'IP")
                    
                    # Verifica che il file contenga la query di test
                    if "INSERT INTO test_table" in content:
                        print(f"  ✅ Il file contiene la query di test")
                    else:
                        print(f"  ❌ Il file non contiene la query di test")
                
                break
        
        if not found:
            print(f"❌ Errore: file per l'ambiente {env_name} non trovato")
    
    print("\nTest completato!")
    return output_files

if __name__ == "__main__":
    output_files = test_save_queries()
    
    # Chiedi all'utente se vuole eliminare i file di test
    response = input("\nVuoi eliminare i file di test generati? (s/n): ")
    if response.lower() == 's':
        for file in output_files:
            try:
                os.remove(file)
                print(f"File eliminato: {os.path.basename(file)}")
            except Exception as e:
                print(f"Errore nell'eliminazione del file {file}: {str(e)}")
        print("Tutti i file di test sono stati eliminati.")
    else:
        print("I file di test non sono stati eliminati.")
