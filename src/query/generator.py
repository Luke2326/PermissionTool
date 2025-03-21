from typing import Dict, List, Optional
import pandas as pd
import logging
from concurrent.futures import ThreadPoolExecutor
import os
from datetime import datetime

from config.constants import OUTPUT_DIR, QUERY_FILE_PREFIX, EXERCISE_TYPE_MAP, ENVIRONMENTS
from src.utils.helpers import ensure_output_directory, generate_timestamp, extract_excel_filename, get_next_version

# Create case-insensitive version of EXERCISE_TYPE_MAP
CASE_INSENSITIVE_EXERCISE_MAP = {k.lower(): v for k, v in EXERCISE_TYPE_MAP.items()}

class QueryGenerator:
    def __init__(self, sheets_data: Dict[str, pd.DataFrame], file_path: str, id: Optional[str] = None):
        """Initialize QueryGenerator with sheets data and metadata."""
        self.sheets_data = sheets_data
        self.file_path = file_path
        self.id = id
        self.queries: List[str] = []
    
    def generate_query(self, sheet_name: str, row: pd.Series) -> Optional[str]:
        """Generate a single query based on sheet type and row data."""
        if pd.isna(row['Delta']):
            return None
            
        query_funcs = {
            'Prometheus Entities': self._generate_entities_query,
            'Prometheus File Types': self._generate_file_types_query,
            'Prometheus Data Items': self._generate_data_items_query,
            'Prometheus Groups': self._generate_groups_query,
            'Prometheus Functionalities': self._generate_functionalities_query,
            'Prometheus Permissions': self._generate_permission_query,
            'Prometheus Roles': self._generate_roles_query,
            'Prometheus Permission Set': self._generate_permission_set_query,
            'Prometheus Set Role Group Ver': self._generate_set_role_group_ver_query,
        }
        
        query_func = query_funcs.get(sheet_name)
        if query_func:
            return query_func(row)
        return None

    def _generate_entities_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            return f"""INSERT INTO public."TBM_Entities" ("Id","Name", "EntityTypeId", "CountryId", "NodeLevelId","ObjectType") 
                    VALUES (
                        (select MAX("Id") + 1 from public."TBM_Entities"),
                        ''{row["Entity"]}'',
                        (SELECT "Id" FROM public."TBM_EntityTypes" WHERE "Name" = ''{row["Entity Type"]}''),
                        (SELECT "Id" FROM public."TBM_Countries" where "Description" = ''{row["Country Code"]}''),
                        (SELECT "Id" FROM public."TBM_NodeLevels" WHERE "Name" = ''{row["Node Level"]}''),
                        1
                    );"""
        return None

    def _generate_data_items_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            return f"""INSERT INTO public."TBM_FileTypeEntities" ("EntityId", "FileTypeId") 
                    VALUES (
                        (SELECT "Id" FROM public."TBM_Entities" WHERE "Name" = ''{row["Entity"]}''),
                        (SELECT "Id" FROM public."TBM_FileTypes" WHERE "Name" = ''{row["File Type"]}'')
                    );"""
        elif row['Delta'].upper() == 'DELETE':
            return f"""DELETE FROM public."TBM_FileTypeEntities" 
                    WHERE "EntityId" = (SELECT "Id" FROM public."TBM_Entities" WHERE "Name" = ''{row["Entity"]}'')
                    AND "FileTypeId" = (SELECT "Id" FROM public."TBM_FileTypes" WHERE "Name" = ''{row["File Type"]}'');"""
        return None

    def _generate_groups_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            exercise_id = CASE_INSENSITIVE_EXERCISE_MAP.get(row["Exercise Type"].lower(), row["Exercise Type"])
            return f"""INSERT INTO public."TBM_Groups" ("Id","Name","Description","LongDescription","ExerciseTypeId","Hidden","IsDeleted")
                    VALUES (
                        (select MAX("Id") + 1 from public."TBM_Groups"),
                        ''{row["Group Code"]}'',
                        ''{row["Description"]}'',
                        NULL,
                        {exercise_id},
                        false,
                        false
                    );
                    INSERT INTO public."TBW_GroupRootNodes" ("GroupId","EntityId","IsDeleted")
                    VALUES (
                        (SELECT "Id" FROM public."TBM_Groups" WHERE "Name" = ''{row["Group Code"]}'' and "ExerciseTypeId"={exercise_id}),
                        (select "Id" as EntityId from public."TBM_Entities" WHERE "Name" = ''{row["Reference Node"]}''),
                        false
                    );"""
        elif row['Delta'].upper() == 'DELETE':
            exercise_id = CASE_INSENSITIVE_EXERCISE_MAP.get(row["Exercise Type"].lower(), row["Exercise Type"])
            return f"""DELETE FROM public."TBW_GroupRootNodes" 
                    WHERE "GroupId" = (SELECT "Id" FROM public."TBM_Groups" WHERE "Name" = ''{row["Group Code"]}'' and "ExerciseTypeId" = {exercise_id});
                    DELETE FROM public."TBW_RolesVersions" 
                    WHERE "GroupId" = (SELECT "Id" FROM public."TBM_Groups" WHERE "Name" = ''{row["Group Code"]}'' and "ExerciseTypeId" = {exercise_id});
                    DELETE FROM public."TBW_NodeVisibilities" 
                    WHERE "GroupId" = (SELECT "Id" FROM public."TBM_Groups" WHERE "Name" = ''{row["Group Code"]}'' and "ExerciseTypeId" = {exercise_id});
                    DELETE FROM public."TBM_Profiles" 
                    WHERE "GroupId" = (SELECT "Id" FROM public."TBM_Groups" WHERE "Name" = ''{row["Group Code"]}'' and "ExerciseTypeId" = {exercise_id});
                    DELETE FROM public."TBM_Groups" 
                    WHERE "Name" = ''{row["Group Code"]}'' and "ExerciseTypeId" = {exercise_id};"""
        return None

    def _generate_permission_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            exercise_id = CASE_INSENSITIVE_EXERCISE_MAP.get(row["Exercise Type"].lower(), row["Exercise Type"])
            file_type_condition = f"(SELECT \"Id\" FROM \"TBM_FileTypes\" WHERE \"Name\" = ''{row['FileType Name']}'')" if pd.notna(row['FileType Name']) else "0"
            entity_condition = f"(SELECT \"Id\" FROM \"TBM_Entities\" WHERE \"Name\" = ''{row['Entity Name']}'')" if pd.notna(row['Entity Name']) else "0"
            
            return f"""INSERT INTO public."TBM_Profiles" ("ExerciseTypeId","FileTypeId","EntityId","PermissionId","GroupId","IsDeleted")
                    VALUES (
                        {exercise_id},
                        {file_type_condition},
                        {entity_condition},
                        (SELECT "Id" FROM "TBM_Permissions" WHERE "Name" = ''{row["Functionality Name"]}''),
                        (SELECT "Id" FROM "TBM_Groups" where "Name" = ''{row["Group Code"]}'' and "ExerciseTypeId" = {exercise_id}),
                        false
                    )
                    ON CONFLICT ("GroupId", "PermissionId", "FileTypeId", "EntityId") DO NOTHING;"""
        elif row['Delta'].upper() == 'DELETE':
            exercise_id = CASE_INSENSITIVE_EXERCISE_MAP.get(row["Exercise Type"].lower(), row["Exercise Type"])
            file_type_condition = f"AND \"FileTypeId\" = (SELECT \"Id\" FROM \"TBM_FileTypes\" WHERE \"Name\" = ''{row['FileType Name']}'')" if pd.notna(row['FileType Name']) else ""
            entity_condition = f"AND \"EntityId\" = (SELECT \"Id\" FROM \"TBM_Entities\" WHERE \"Name\" = ''{row['Entity Name']}'')" if pd.notna(row['Entity Name']) else ""
            
            return f"""DELETE FROM public."TBM_Profiles" 
                    WHERE "GroupId" = (SELECT "Id" FROM "TBM_Groups" where "Name" = ''{row["Group Code"]}'' and "ExerciseTypeId" = {exercise_id})
                    {entity_condition}
                    {file_type_condition}
                    AND "PermissionId" = (SELECT "Id" FROM "TBM_Permissions" WHERE "Name" = ''{row["Functionality Name"]}'');"""
        return None

    def _generate_roles_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            return f"""INSERT INTO public."TBM_Roles" ("Name","Description")
                    VALUES (''{row["Role Unique Name"]}'',''{row["Role Name"]}'');"""
        elif row['Delta'].upper() == 'DELETE':
            return f"""DELETE FROM public."TBM_Roles" WHERE "Name" = ''{row["Role Unique Name"]}'';"""
        return None

    def _generate_permission_set_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            exercise_id = CASE_INSENSITIVE_EXERCISE_MAP.get(row["Exercise Type"].lower(), row["Exercise Type"])
            return f"""INSERT INTO public."TBM_ProfileSet" ("Name","ExerciseTypeId")
                    VALUES (''{row["ProfileSetName"]}'',{exercise_id});
                    INSERT INTO public."TBM_ProfileSetVersion" ("ProfileSetId","Name","Description","Version")
                    VALUES (
                        (select "Id" from public."TBM_ProfileSet" WHERE "Name" = ''{row["ProfileSetName"]}''),
                        ''{row["ProfileSetVersionName"]}'',
                        ''{row["Set Version Name"]}'',
                        ''{row["Version Number"]}''
                    );"""
        elif row['Delta'].upper() == 'DELETE':
            return f"""DELETE FROM public."TBM_ProfileSetVersion" 
                    WHERE "Description" = ''{row["Set Version Name"]}'';
                    DELETE FROM public."TBM_ProfileSet"
                    WHERE "Name" = ''{row["ProfileSetName"]}'';"""
        return None

    def _generate_set_role_group_ver_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            exercise_id = CASE_INSENSITIVE_EXERCISE_MAP.get(row["Exercise Type"].lower(), row["Exercise Type"])
            return f"""INSERT INTO public."TBW_RolesVersions" ("ProfileSetId","ProfileSetVersionId","RoleId","GroupId","IsDeleted")
                    VALUES (
                        (select "ProfileSetId" from public."TBM_ProfileSetVersion" where "Description" = ''{row["Set Version Name"]}''),
                        (select "Id" from public."TBM_ProfileSetVersion" where "Description" = ''{row["Set Version Name"]}''),
                        (select "Id" from public."TBM_Roles" WHERE "Name" = ''{row["Role Unique Name"]}''),
                        (select "Id" from public."TBM_Groups" WHERE "Name" = ''{row["Group Unique Name"]}'' and "ExerciseTypeId" = {exercise_id}),
                        false
                    )
                    ON CONFLICT ("ProfileSetId", "ProfileSetVersionId", "RoleId", "GroupId") DO NOTHING;"""
        elif row['Delta'].upper() == 'DELETE':
            return f"""UPDATE public."TBW_RolesVersions"
                    SET "IsDeleted" = true
                    WHERE "ProfileSetVersionId" = (select "Id" from public."TBM_ProfileSetVersion" where "Description" = ''{row["Set Version Name"]}'')
                    AND "RoleId" = (select "Id" from public."TBM_Roles" WHERE "Name" = ''{row["Role Unique Name"]}'')
                    AND "GroupId" = (select "Id" from public."TBM_Groups" WHERE "Name" = ''{row["Group Unique Name"]}'');"""
        return None

    def _generate_file_types_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            mandatory = 'Y' if pd.notna(row['Mandatory']) and row['Mandatory'].upper() == 'Y' else 'N'
            mandatoryFormatted = 'true' if mandatory == 'Y' else 'false'
            formatted_value = row["ValidFileExtension"].replace('|', ',')

            return f"""-- DA INSERIRE MANUALMENTE ID
                    INSERT INTO public."TBM_FileTypes" ("Id","Name","ClassificationId","DataSourceId","RiskModuleId","ValidFileExtension","IgnoreRegex","Mandatory","DataQualityTopic","PushableAsDataitem","PushableAsTemplate", "PushableAsEngineOutput", "TrackingChanges")
                    VALUES (
                        'CONTROLLARE FILE CENSIMENTI'
                        ''{row["Name"]}'',
                        (SELECT "Id" FROM public."TBM_Classifications" WHERE "Name" = ''{row["Classification"]}''),
                        (SELECT "Id" from public."TBM_DataSources" WHERE "Name" = ''{row["DataSource"]}''),
                        (SELECT "Id" from public."TBM_RiskModules" WHERE "Name" = ''{row["RiskModule"]}''),
                        ''{formatted_value}'',
                        ''n$|N$'',
                        {mandatoryFormatted},
                        ''dataqualitydataitemcheck'',
                        3,
                        0,
                        3,
                        0
                    );"""
        elif row['Delta'].upper() == 'DELETE':
            return f"""DELETE FROM public."TBM_FileTypes" WHERE "Name" = ''{row["Name"]}'';"""
        return None

    def _generate_functionalities_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            data_items_required = row.get('DataItemsRequired', 'N')
            data_items_required = 'Y' if pd.notna(data_items_required) and data_items_required.upper() == 'Y' else 'N'
            value = 'true' if data_items_required == 'Y' else 'false'
            
            return f"""INSERT INTO public."TBM_Permissions" ("Name", "DataItemRequired", "Description", "Id", "Versionable") 
                    VALUES (
                        ''{row["Name"]}'',
                        {value},
                        ''{row["Description"]}'',
                        (SELECT MAX("Id") + 1 FROM public."TBM_Permissions"),
                        false
                    );"""
        elif row['Delta'].upper() == 'DELETE':
            return f"""DELETE FROM public."TBM_Permissions" WHERE "Name" = ''{row["Name"]}'';"""
        return None

    def generate_all_queries(self) -> bool:
        """Generate all queries in parallel."""
        has_errors = False
        with ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
            # Manteniamo traccia dello sheet_name insieme al future
            futures_with_context = []
            for sheet_name, df in self.sheets_data.items():
                for idx, row in df.iterrows():
                    future = executor.submit(self.generate_query, sheet_name, row)
                    futures_with_context.append((future, sheet_name, idx))
            
            for future, sheet_name, idx in futures_with_context:
                try:
                    query = future.result()
                    if query:
                        self.queries.append(query)
                except KeyError as e:
                    has_errors = True
                    logging.error(f"Errore nella generazione della query - Sheet: {sheet_name}, Riga: {idx + 2}, Colonna mancante: {str(e)}")
                except Exception as e:
                    has_errors = True
                    logging.error(f"Errore generico nella generazione della query - Sheet: {sheet_name}, Riga: {idx + 2}: {str(e)}")
            
        return not has_errors

    def save_queries(self, executed_by: Optional[str] = None) -> List[str]:
        """
        Save generated queries to multiple files, one for each environment.
        
        Args:
            executed_by: Nome della persona che esegue lo script
            
        Returns:
            List[str]: List of paths to the generated SQL files
        """
        if not self.queries:
            logging.warning("Nessuna query da salvare.")
            return []
        
        output_dir = ensure_output_directory(OUTPUT_DIR)
        date_str = generate_timestamp()
        
        # Estrai il nome del file Excel senza estensione
        excel_filename = extract_excel_filename(self.file_path)
        
        # Rimuovi l'ID dal nome del file se è già presente nel formato [ID]
        if self.id and f"[{self.id}]" in excel_filename:
            # L'ID è già nel nome del file, usa il nome così com'è
            base_filename = excel_filename
        elif self.id:
            # Aggiungi l'ID al nome del file
            base_filename = f"[{self.id}] {excel_filename}"
        else:
            # Nessun ID disponibile
            base_filename = excel_filename
        
        output_files = []
        
        # Genera un file per ogni ambiente
        for env_name, env_config in ENVIRONMENTS.items():
            # Ottieni l'indirizzo IP dell'ambiente (primo host trovato)
            env_ip = env_config[0]["config"]["host"]
            
            # Format queries with transaction control
            formatted_output = []
            formatted_output.extend([
                "-- Script generato automaticamente",
                f"-- Data generazione: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                f"-- File di origine: {os.path.basename(self.file_path)}",
                f"-- Ambiente: {env_name}",
                "",
                "DO $$",
                "DECLARE",
                "    current_ip TEXT;",
                "    expected_ip TEXT := '" + env_ip + "';",
                "BEGIN",
                "    -- Ottieni l'IP del server corrente",
                "    SELECT inet_server_addr()::TEXT INTO current_ip;",
                "",
                "    -- Verifica che l'IP corrisponda all'ambiente designato",
                "    IF NOT current_ip LIKE expected_ip || '%' THEN",
                "        RAISE EXCEPTION 'Errore: Lo script non sta venendo eseguito nell''ambiente corretto: " + env_name + " (%).',",
                "                        expected_ip;",
                "    END IF;",
                "END $$;",
                "",
                "BEGIN;",
                "",
                "-- Variabili per il tracciamento dell'esecuzione",
                "DO $$",
                "DECLARE",
                "    v_start_time TIMESTAMP := clock_timestamp();",
                "    v_operation_count INT := 0;",
                "    v_log_id INT;",
                "    v_file_name TEXT := '" + os.path.basename(excel_filename) + "';",
                "    v_executed_by TEXT := '" + (executed_by or 'Unknown') + "';",
                "    v_execution_duration INTERVAL;",
                "    operazioni TEXT[] := ARRAY["
            ])
            
            # Formatta le query come array di stringhe
            for i, query in enumerate(self.queries):
                if i < len(self.queries) - 1:
                    formatted_output.append(f"        '{query}',")
                else:
                    formatted_output.append(f"        '{query}'")
            
            formatted_output.extend([
                "    ];",
                "    i INTEGER;",
                "    query_in_esecuzione TEXT;",
                "BEGIN",
                "    -- Esegui le operazioni",
                "    FOR i IN 1..array_length(operazioni, 1) LOOP",
                "        BEGIN",
                "            query_in_esecuzione := operazioni[i];",
                "            EXECUTE query_in_esecuzione;",
                "            v_operation_count := v_operation_count + 1;",
                "        EXCEPTION",
                "            WHEN OTHERS THEN",
                "                RAISE NOTICE 'Errore nell''operazione %: %', i, SQLERRM;",
                "                RAISE NOTICE 'Query che ha causato l''errore: %', query_in_esecuzione;",
                "        END;",
                "    END LOOP;",
                "",
                "    -- Calcola la durata dell'esecuzione",
                "    v_execution_duration := clock_timestamp() - v_start_time;",
                "",
                "    -- Registra le informazioni di esecuzione nella tabella TBW_ServiceRequestLog",
                "    -- Questa parte viene eseguita solo quando si fa COMMIT",
                "    IF EXISTS (SELECT 1 FROM pg_tables WHERE tablename = 'TBW_ServiceRequestLog') THEN",
                "        INSERT INTO public.\"TBW_ServiceRequestLog\" (\"FileName\", \"ExecutedBy\", \"ExecutionDate\", \"ExecutionDuration\", \"OperationCount\")",
                "        VALUES (v_file_name, v_executed_by, clock_timestamp(), v_execution_duration, v_operation_count)",
                "        RETURNING \"Id\" INTO v_log_id;",
                "",
                "        RAISE NOTICE 'Esecuzione registrata con ID: %', v_log_id;",
                "        RAISE NOTICE 'Operazioni eseguite: %', v_operation_count;",
                "        RAISE NOTICE 'Durata esecuzione: %', v_execution_duration;",
                "    ELSE",
                "        RAISE NOTICE 'Tabella TBW_ServiceRequestLog non trovata. Le informazioni di esecuzione non sono state registrate.';",
                "    END IF;",
                "END $$;",
                "",
                "ROLLBACK;",
                "-- Per eseguire le operazioni, rimuovere il commento dalla riga seguente e commentare la riga ROLLBACK sopra",
                "--COMMIT;"
            ])
            
            # Ottieni la prossima versione disponibile
            env_base_filename = f"[{self.id}][{env_name}] {excel_filename}" if self.id else f"[{env_name}] {excel_filename}"
            version = get_next_version(output_dir, env_base_filename, date_str)
            
            # Crea il nome completo del file
            filename = f"{env_base_filename} {date_str} {version}.sql"
            output_file = os.path.join(output_dir, filename)
            
            with open(output_file, 'w') as f:
                f.write("\n".join(formatted_output))
            
            output_files.append(output_file)
            logging.info(f"File SQL generato per l'ambiente {env_name}: {output_file}")
        
        return output_files
