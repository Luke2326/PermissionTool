from typing import Dict, List, Optional
import pandas as pd
import logging
from concurrent.futures import ThreadPoolExecutor
import os

from config.constants import OUTPUT_DIR, QUERY_FILE_PREFIX, EXERCISE_TYPE_MAP
from src.utils.helpers import ensure_output_directory, generate_timestamp

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
            return f"""WITH entity_data AS (
                        SELECT MAX("Id") + 1 as new_id FROM public."TBM_Entities"
                    ),
                    entity_type AS (
                        SELECT "Id" FROM public."TBM_EntityTypes" WHERE "Name" = ''{row["Entity Type"]}''
                    ),
                    country AS (
                        SELECT "Id" FROM public."TBM_Countries" where "Description" = ''{row["Country Code"]}''
                    ),
                    node_level AS (
                        SELECT "Id" FROM public."TBM_NodeLevels" WHERE "Name" = ''{row["Node Level"]}''
                    )
                    INSERT INTO public."TBM_Entities" ("Id","Name", "EntityTypeId", "CountryId", "NodeLevelId","ObjectType") 
                    SELECT 
                        entity_data.new_id,
                        ''{row["Entity"]}'',
                        entity_type."Id",
                        country."Id",
                        node_level."Id",
                        1
                    FROM 
                        entity_data, entity_type, country, node_level;"""
        return None

    def _generate_data_items_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            return f"""WITH entity AS (
                        SELECT "Id" FROM public."TBM_Entities" WHERE "Name" = ''{row["Entity"]}''
                    ),
                    file_type AS (
                        SELECT "Id" FROM public."TBM_FileTypes" WHERE "Name" = ''{row["File Type"]}''
                    )
                    INSERT INTO public."TBM_FileTypeEntities" ("EntityId", "FileTypeId") 
                    SELECT entity."Id", file_type."Id"
                    FROM entity, file_type;"""
        elif row['Delta'].upper() == 'DELETE':
            return f"""WITH entity AS (
                        SELECT "Id" FROM public."TBM_Entities" WHERE "Name" = ''{row["Entity"]}''
                    ),
                    file_type AS (
                        SELECT "Id" FROM public."TBM_FileTypes" WHERE "Name" = ''{row["File Type"]}''
                    )
                    DELETE FROM public."TBM_FileTypeEntities" 
                    WHERE "EntityId" IN (SELECT "Id" FROM entity)
                    AND "FileTypeId" IN (SELECT "Id" FROM file_type);"""
        return None

    def _generate_groups_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            exercise_id = CASE_INSENSITIVE_EXERCISE_MAP.get(row["Exercise Type"].lower(), row["Exercise Type"])
            return f"""WITH new_group AS (
                        INSERT INTO public."TBM_Groups" ("Id","Name","Description","LongDescription","ExerciseTypeId","Hidden","IsDeleted")
                        VALUES (
                            (select MAX("Id") + 1 from public."TBM_Groups"),
                            ''{row["Group Code"]}'',
                            ''{row["Description"]}'',
                            NULL,
                            {exercise_id},
                            false,
                            false
                        )
                        RETURNING "Id"
                    ),
                    entity AS (
                        SELECT "Id" FROM public."TBM_Entities" WHERE "Name" = ''{row["Reference Node"]}''
                    )
                    INSERT INTO public."TBW_GroupRootNodes" ("GroupId","EntityId","IsDeleted")
                    SELECT new_group."Id", entity."Id", false
                    FROM new_group, entity;"""
        elif row['Delta'].upper() == 'DELETE':
            exercise_id = CASE_INSENSITIVE_EXERCISE_MAP.get(row["Exercise Type"].lower(), row["Exercise Type"])
            return f"""WITH target_group AS (
                        SELECT "Id" FROM public."TBM_Groups" 
                        WHERE "Name" = ''{row["Group Code"]}'' AND "ExerciseTypeId" = {exercise_id}
                    )
                    DELETE FROM public."TBW_GroupRootNodes" WHERE "GroupId" IN (SELECT "Id" FROM target_group);
                    DELETE FROM public."TBW_RolesVersions" WHERE "GroupId" IN (SELECT "Id" FROM target_group);
                    DELETE FROM public."TBW_NodeVisibilities" WHERE "GroupId" IN (SELECT "Id" FROM target_group);
                    DELETE FROM public."TBM_Profiles" WHERE "GroupId" IN (SELECT "Id" FROM target_group);
                    DELETE FROM public."TBM_Groups" WHERE "Id" IN (SELECT "Id" FROM target_group);"""
        return None

    def _generate_permission_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            exercise_id = CASE_INSENSITIVE_EXERCISE_MAP.get(row["Exercise Type"].lower(), row["Exercise Type"])
            
            cte_parts = ["WITH"]
            select_parts = []
            from_parts = []
            
            # Always include permission and group CTEs
            cte_parts.append(f"""
                permission AS (
                    SELECT "Id" FROM "TBM_Permissions" WHERE "Name" = ''{row["Functionality Name"]}''
                ),
                group_data AS (
                    SELECT "Id" FROM "TBM_Groups" where "Name" = ''{row["Group Code"]}'' and "ExerciseTypeId" = {exercise_id}
                )""")
            
            select_parts.extend([f"{exercise_id}", "permission.\"Id\"", "group_data.\"Id\""])
            from_parts.extend(["permission", "group_data"])
            
            # Conditionally include file_type and entity CTEs
            if pd.notna(row['FileType Name']):
                cte_parts.append(f"""
                    file_type AS (
                        SELECT "Id" FROM "TBM_FileTypes" WHERE "Name" = ''{row['FileType Name']}''
                    )""")
                select_parts.append("file_type.\"Id\"")
                from_parts.append("file_type")
            else:
                select_parts.append("0")
            
            if pd.notna(row['Entity Name']):
                cte_parts.append(f"""
                    entity AS (
                        SELECT "Id" FROM "TBM_Entities" WHERE "Name" = ''{row['Entity Name']}''
                    )""")
                select_parts.append("entity.\"Id\"")
                from_parts.append("entity")
            else:
                select_parts.append("0")
            
            # Combine all parts into the final query
            final_cte = " ".join(cte_parts)
            final_select = ", ".join(select_parts)
            final_from = ", ".join(from_parts)
            
            return f"""{final_cte}
                    INSERT INTO public."TBM_Profiles" ("ExerciseTypeId","PermissionId","GroupId","FileTypeId","EntityId","IsDeleted")
                    SELECT {final_select}, false
                    FROM {final_from}
                    ON CONFLICT ("GroupId", "PermissionId", "FileTypeId", "EntityId") DO NOTHING;"""
        elif row['Delta'].upper() == 'DELETE':
            exercise_id = CASE_INSENSITIVE_EXERCISE_MAP.get(row["Exercise Type"].lower(), row["Exercise Type"])
            
            cte_parts = ["WITH"]
            where_conditions = [
                f"\"ExerciseTypeId\" = {exercise_id}",
                "\"PermissionId\" IN (SELECT \"Id\" FROM permission)",
                "\"GroupId\" IN (SELECT \"Id\" FROM group_data)"
            ]
            
            # Always include permission and group CTEs
            cte_parts.append(f"""
                permission AS (
                    SELECT "Id" FROM "TBM_Permissions" WHERE "Name" = ''{row["Functionality Name"]}''
                ),
                group_data AS (
                    SELECT "Id" FROM "TBM_Groups" where "Name" = ''{row["Group Code"]}'' and "ExerciseTypeId" = {exercise_id}
                )""")
            
            # Conditionally include file_type and entity CTEs and conditions
            if pd.notna(row['FileType Name']):
                cte_parts.append(f"""
                    file_type AS (
                        SELECT "Id" FROM "TBM_FileTypes" WHERE "Name" = ''{row['FileType Name']}''
                    )""")
                where_conditions.append("\"FileTypeId\" IN (SELECT \"Id\" FROM file_type)")
            
            if pd.notna(row['Entity Name']):
                cte_parts.append(f"""
                    entity AS (
                        SELECT "Id" FROM "TBM_Entities" WHERE "Name" = ''{row['Entity Name']}''
                    )""")
                where_conditions.append("\"EntityId\" IN (SELECT \"Id\" FROM entity)")
            
            # Combine all parts into the final query
            final_cte = " ".join(cte_parts)
            final_where = " AND ".join(where_conditions)
            
            return f"""{final_cte}
                    DELETE FROM public."TBM_Profiles" 
                    WHERE {final_where};"""
        return None

    def _generate_roles_query(self, row: pd.Series) -> Optional[str]:
        # This query is already simple, but we'll keep the pattern consistent
        if row['Delta'].upper() == 'INSERT':
            return f"""INSERT INTO public."TBM_Roles" ("Name","Description")
                    VALUES (''{row["Role Unique Name"]}'',''{row["Role Name"]}'');"""
        elif row['Delta'].upper() == 'DELETE':
            return f"""DELETE FROM public."TBM_Roles" WHERE "Name" = ''{row["Role Unique Name"]}'';"""
        return None

    def _generate_permission_set_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            exercise_id = CASE_INSENSITIVE_EXERCISE_MAP.get(row["Exercise Type"].lower(), row["Exercise Type"])
            return f"""WITH new_profile_set AS (
                        INSERT INTO public."TBM_ProfileSet" ("Name","ExerciseTypeId")
                        VALUES (''{row["ProfileSetName"]}'',{exercise_id})
                        RETURNING "Id"
                    )
                    INSERT INTO public."TBM_ProfileSetVersion" ("ProfileSetId","Name","Description","Version")
                    SELECT 
                        new_profile_set."Id",
                        ''{row["ProfileSetVersionName"]}'',
                        ''{row["Set Version Name"]}'',
                        ''{row["Version Number"]}''
                    FROM new_profile_set;"""
        elif row['Delta'].upper() == 'DELETE':
            return f"""WITH profile_set_version AS (
                        SELECT "ProfileSetId" 
                        FROM public."TBM_ProfileSetVersion" 
                        WHERE "Description" = ''{row["Set Version Name"]}''
                    )
                    DELETE FROM public."TBM_ProfileSetVersion" 
                    WHERE "Description" = ''{row["Set Version Name"]}'';
                    
                    DELETE FROM public."TBM_ProfileSet"
                    WHERE "Id" IN (SELECT "ProfileSetId" FROM profile_set_version);"""
        return None

    def _generate_set_role_group_ver_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            exercise_id = CASE_INSENSITIVE_EXERCISE_MAP.get(row["Exercise Type"].lower(), row["Exercise Type"])
            return f"""WITH profile_set_version AS (
                        SELECT "Id", "ProfileSetId"
                        FROM public."TBM_ProfileSetVersion" 
                        WHERE "Description" = ''{row["Set Version Name"]}''
                    ),
                    role_data AS (
                        SELECT "Id" 
                        FROM public."TBM_Roles" 
                        WHERE "Name" = ''{row["Role Unique Name"]}''
                    ),
                    group_data AS (
                        SELECT "Id" 
                        FROM public."TBM_Groups" 
                        WHERE "Name" = ''{row["Group Unique Name"]}'' AND "ExerciseTypeId" = {exercise_id}
                    )
                    INSERT INTO public."TBW_RolesVersions" ("ProfileSetId","ProfileSetVersionId","RoleId","GroupId","IsDeleted")
                    SELECT 
                        profile_set_version."ProfileSetId",
                        profile_set_version."Id",
                        role_data."Id",
                        group_data."Id",
                        false
                    FROM 
                        profile_set_version, role_data, group_data
                    ON CONFLICT ("ProfileSetId", "ProfileSetVersionId", "RoleId", "GroupId") DO NOTHING;"""
        elif row['Delta'].upper() == 'DELETE':
            return f"""WITH profile_set_version AS (
                        SELECT "Id" 
                        FROM public."TBM_ProfileSetVersion" 
                        WHERE "Description" = ''{row["Set Version Name"]}''
                    ),
                    role_data AS (
                        SELECT "Id" 
                        FROM public."TBM_Roles" 
                        WHERE "Name" = ''{row["Role Unique Name"]}''
                    ),
                    group_data AS (
                        SELECT "Id" 
                        FROM public."TBM_Groups" 
                        WHERE "Name" = ''{row["Group Unique Name"]}''
                    )
                    UPDATE public."TBW_RolesVersions"
                    SET "IsDeleted" = true
                    WHERE "ProfileSetVersionId" IN (SELECT "Id" FROM profile_set_version)
                    AND "RoleId" IN (SELECT "Id" FROM role_data)
                    AND "GroupId" IN (SELECT "Id" FROM group_data);"""
        return None

    def _generate_file_types_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            mandatory = 'Y' if pd.notna(row['Mandatory']) and row['Mandatory'].upper() == 'Y' else 'N'
            mandatoryFormatted = 'true' if mandatory == 'Y' else 'false'
            formatted_value = row["ValidFileExtension"].replace('|', ',')

            return f"""WITH classification AS (
                        SELECT "Id" FROM public."TBM_Classifications" WHERE "Name" = ''{row["Classification"]}''
                    ),
                    data_source AS (
                        SELECT "Id" FROM public."TBM_DataSources" WHERE "Name" = ''{row["DataSource"]}''
                    ),
                    risk_module AS (
                        SELECT "Id" FROM public."TBM_RiskModules" WHERE "Name" = ''{row["RiskModule"]}''
                    )
                    -- DA INSERIRE MANUALMENTE ID
                    INSERT INTO public."TBM_FileTypes" ("Id","Name","ClassificationId","DataSourceId","RiskModuleId","ValidFileExtension","IgnoreRegex","Mandatory","DataQualityTopic","PushableAsDataitem","PushableAsTemplate", "PushableAsEngineOutput", "TrackingChanges")
                    SELECT
                        'CONTROLLARE FILE CENSIMENTI',
                        ''{row["Name"]}'',
                        classification."Id",
                        data_source."Id",
                        risk_module."Id",
                        ''{formatted_value}'',
                        ''n$|N$'',
                        {mandatoryFormatted},
                        ''dataqualitydataitemcheck'',
                        3,
                        0,
                        3,
                        0
                    FROM
                        classification, data_source, risk_module;"""
        elif row['Delta'].upper() == 'DELETE':
            return f"""DELETE FROM public."TBM_FileTypes" WHERE "Name" = ''{row["Name"]}'';"""
        return None

    def _generate_functionalities_query(self, row: pd.Series) -> Optional[str]:
        if row['Delta'].upper() == 'INSERT':
            data_items_required = row.get('DataItemsRequired', 'N')
            data_items_required = 'Y' if pd.notna(data_items_required) and data_items_required.upper() == 'Y' else 'N'
            value = 'true' if data_items_required == 'Y' else 'false'
            
            return f"""WITH max_id AS (
                        SELECT MAX("Id") + 1 AS new_id FROM public."TBM_Permissions"
                    )
                    INSERT INTO public."TBM_Permissions" ("Name", "DataItemRequired", "Description", "Id", "Versionable") 
                    SELECT
                        ''{row["Name"]}'',
                        {value},
                        ''{row["Description"]}'',
                        max_id.new_id,
                        false
                    FROM max_id;"""
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
