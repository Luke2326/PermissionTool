from typing import Dict, List, Optional
import pandas as pd
import logging
from concurrent.futures import ThreadPoolExecutor
import os

from config.constants import OUTPUT_DIR, QUERY_FILE_PREFIX, EXERCISE_TYPE_MAP
from src.utils.helpers import ensure_output_directory, generate_timestamp

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
            'Prometheus Data Items': self._generate_data_items_query,
            'Prometheus Groups': self._generate_groups_query,
            'Prometheus Functionalities': self._generate_functionalities_query,
            'Prometheus Permissions': self._generate_permission_query,
            'Prometheus Roles': self._generate_roles_query,
            'Prometheus Permission Set': self._generate_permission_set_query,
            'Prometheus Set Role Group Ver': self._generate_set_role_group_ver_query,
            'Prometheus File Types': self._generate_file_types_query
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
            exercise_id = EXERCISE_TYPE_MAP.get(row["Exercise Type"], row["Exercise Type"])
            return f"""INSERT INTO public."TBM_Groups" ("Id","Name","Description","LongDescription","ExerciseTypeId")
                    VALUES (
                        (select MAX("Id") + 1 from public."TBM_Groups"),
                        ''{row["Group Code"]}'',
                        ''{row["Description"]}'',
                        NULL,
                        {exercise_id}
                    );
                    INSERT INTO public."TBW_GroupRootNodes" ("GroupId","EntityId")
                    VALUES (
                        (SELECT "Id" FROM public."TBM_Groups" WHERE "Name" = ''{row["Group Code"]}'' and "ExerciseTypeId"={exercise_id}),
                        (select "Id" as EntityId from public."TBM_Entities" WHERE "Name" = ''{row["Reference Node"]}'')
                    );"""
        elif row['Delta'].upper() == 'DELETE':
            exercise_id = EXERCISE_TYPE_MAP.get(row["Exercise Type"], row["Exercise Type"])
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
            exercise_id = EXERCISE_TYPE_MAP.get(row["Exercise Type"], row["Exercise Type"])
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
            exercise_id = EXERCISE_TYPE_MAP.get(row["Exercise Type"], row["Exercise Type"])
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
            exercise_id = EXERCISE_TYPE_MAP.get(row["Exercise Type"], row["Exercise Type"])
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
            exercise_id = EXERCISE_TYPE_MAP.get(row["Exercise Type"], row["Exercise Type"])
            return f"""INSERT INTO public."TBW_RolesVersions" ("ProfileSetId","ProfileSetVersionId","RoleId","GroupId")
                    VALUES (
                        (select "ProfileSetId" from public."TBM_ProfileSetVersion" where "Description" = ''{row["Set Version Name"]}''),
                        (select "Id" from public."TBM_ProfileSetVersion" where "Description" = ''{row["Set Version Name"]}''),
                        (select "Id" from public."TBM_Roles" WHERE "Name" = ''{row["Role Unique Name"]}''),
                        (select "Id" from public."TBM_Groups" WHERE "Name" = ''{row["Group Unique Name"]}'' and "ExerciseTypeId" = {exercise_id})
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
            formatted_value = row["Format"].replace('|', ',')

            return f"""INSERT INTO public."TBM_FileTypes" ("Name","ClassificationId","DataSourceId","ValidFileExtension","Mandatory","RiskModuleId","IgnoreRegex")
                    VALUES (
                        '{row["File Type"]}',
                        (SELECT "Id" FROM public."TBM_Classifications" WHERE "Name" = ''{row["Classification"]}''),
                        (SELECT "Id" from public."TBM_DataSources" WHERE "Name" = ''{row["Data Source"]}''),
                        '{formatted_value}',
                        {mandatoryFormatted},
                        (SELECT "Id" from public."TBM_RiskModules" WHERE "Name" = ''{row["Risk Module"]}''),
                        'n$|N$'
                    );"""
        elif row['Delta'].upper() == 'DELETE':
            return f"""DELETE FROM public."TBM_FileTypes" WHERE "Name" = ''{row["File Type"]}'';"""
        return None

    def _generate_functionalities_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            data_items_required = row.get('Data Items Required', 'N')
            data_items_required = 'Y' if pd.notna(data_items_required) and data_items_required.upper() == 'Y' else 'N'
            value = 'true' if data_items_required == 'Y' else 'false'
            
            return f"""INSERT INTO public."TBM_Permissions" ("Name", "DataItemRequired", "Description", "Id", "Versionable") 
                    VALUES (
                        ''{row["Functionality"]}'',
                        {value},
                        ''{row["Description"]}'',
                        (SELECT MAX("Id") + 1 FROM public."TBM_Permissions"),
                        false
                    );"""
        elif row['Delta'].upper() == 'DELETE':
            return f"""DELETE FROM public."TBM_Permissions" WHERE "Name" = ''{row["Functionality"]}'';"""
        return None

    def generate_all_queries(self) -> bool:
        """Generate all queries in parallel."""
        has_errors = False
        with ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
            futures = []
            for sheet_name, df in self.sheets_data.items():
                for idx, row in df.iterrows():
                    futures.append(
                        executor.submit(self.generate_query, sheet_name, row)
                    )
            
            for future in futures:
                try:
                    query = future.result()
                    if query:
                        self.queries.append(query)
                except KeyError as e:
                    has_errors = True
                    logging.error(f"Errore nella generazione della query - colonna mancante: {str(e)}")
                except Exception as e:
                    has_errors = True
                    logging.error(f"Errore nella generazione della query: {str(e)}")
            
        return not has_errors

    def save_queries(self) -> Optional[str]:
        """Save generated queries to a file."""
        if not self.queries:
            logging.warning("Nessuna query da salvare")
            return None
            
        formatted_output = [
            f"--Query generate da {self.file_path}",
            "BEGIN;",
            "",
            "DO $$",
            "DECLARE",
            "    operazioni TEXT[] := ARRAY["
        ]
        
        for i, query in enumerate(self.queries):
            if i < len(self.queries) - 1:
                formatted_output.append(f"        '{query}',\n")
            else:
                formatted_output.append(f"        '{query}'")
        
        formatted_output.extend([
            "    ];",
            "    i INTEGER;",
            "    query_in_esecuzione TEXT;",
            "BEGIN",
            "    FOR i IN 1..array_length(operazioni, 1) LOOP",
            "        BEGIN",
            "            query_in_esecuzione := operazioni[i];",
            "            EXECUTE query_in_esecuzione;",
            "        EXCEPTION",
            "            WHEN OTHERS THEN",
            "                RAISE NOTICE 'Errore nell''operazione %: %', i, SQLERRM;",
            "                RAISE NOTICE 'Query che ha causato l''errore: %', query_in_esecuzione;",
            "        END;",
            "    END LOOP;",
            "END $$;",
            "",
            "ROLLBACK;",
            "--COMMIT;"
        ])
        
        output_dir = ensure_output_directory(OUTPUT_DIR)
        timestamp = generate_timestamp()
        filename = f'{QUERY_FILE_PREFIX}_{self.id}_{timestamp}.sql' if self.id else f'{QUERY_FILE_PREFIX}_{timestamp}.sql'
        output_file = os.path.join(output_dir, filename)
        
        with open(output_file, 'w') as f:
            f.write("\n".join(formatted_output))
        
        return output_file
