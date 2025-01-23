from typing import List, Dict

# Excel sheet names
REQUIRED_SHEETS: List[str] = [
    'Prometheus Entities',
    'Prometheus Data Items',
    'Prometheus Groups',
    'Prometheus Functionalities',
    'Prometheus Permissions',
    'Prometheus Roles',
    'Prometheus Permission Set',
    'Prometheus Set Role Group Ver',
    'Prometheus File Types'
]

# Exercise type mapping
EXERCISE_TYPE_MAP: Dict[str, int] = {
    "Climate Change": 14,
    "Climate Change Investments": 17,
    "Correlation Matrix": 3,
    "ECL Catalogue": 15,
    "Expected Credit Loss": 16,
    "Financial and Credit Asset Pricing": 12,
    "Global": 0,
    "Hierarchy": 4,
    "Internal Model": 1,
    "Internal Model Roll Forward": 6,
    "Loss Data Collection": 13,
    "Loss Function Fitting": 5,
    "Op Risk Calibration": 8,
    "OpRisk Catalogues": 9,
    "OpRisk File Archive": -2,
    "ORA and FSA": 11,
    "Reporting": -1,
    "Standard Formula": 2,
    "Standard Formula Roll Forward": 7,
    "Third Party Risk": 10
}

# Excel configuration
EXCEL_ENGINE = 'pyxlsb'
DEFAULT_DTYPE = {
    'Delta': 'object'
}

# Output configuration
OUTPUT_DIR = "GeneratedQueries"
QUERY_FILE_PREFIX = "generated_queries"

# Environment configurations
ENVIRONMENTS = {
    "SIT": [
        {
            "config": {
                "host": "172.30.2.58",
                "port": "5432",
                "dbname": "DB_PROMETHEUS_DOMAIN",
                "user": "prometheus-admin",
                "password": "7G@},P=eoJk&Ex]i"
            },
            "views": ["Prometheus Entities", "Prometheus Data Items"]
        },
        {
            "config": {
                "host": "172.30.2.58",
                "port": "5432",
                "dbname": "DB_PROMETHEUS_AUTH",
                "user": "prometheus-admin",
                "password": "7G@},P=eoJk&Ex]i"
            },
            "views": ["Prometheus Groups", "Prometheus Permissions", "Prometheus Roles", 
                     "Prometheus Permission Set", "Prometheus Set Role Group Ver"]
        }
    ],
    "UAT": [
        {
            "config": {
                "host": "172.30.2.115",
                "port": "5432",
                "dbname": "DB_PROMETHEUS_DOMAIN",
                "user": "prometheus-admin",
                "password": "Y!osVrO*WdDS]Mn1"
            },
            "views": ["Prometheus Entities", "Prometheus Data Items"]
        },
        {
            "config": {
                "host": "172.30.2.115",
                "port": "5432",
                "dbname": "DB_PROMETHEUS_AUTH",
                "user": "prometheus-admin",
                "password": "Y!osVrO*WdDS]Mn1"
            },
            "views": ["Prometheus Groups", "Prometheus Permissions", "Prometheus Roles", 
                     "Prometheus Permission Set", "Prometheus Set Role Group Ver"]
        }
    ],
    "PREPROD": [
        {
            "config": {
                "host": "172.30.2.95",
                "port": "5432",
                "dbname": "DB_PROMETHEUS_DOMAIN",
                "user": "prometheus-admin",
                "password": "L>XG9T=NhVTdx9aw"
            },
            "views": ["Prometheus Entities", "Prometheus Data Items"]
        },
        {
            "config": {
                "host": "172.30.2.95",
                "port": "5432",
                "dbname": "DB_PROMETHEUS_AUTH",
                "user": "prometheus-admin",
                "password": "L>XG9T=NhVTdx9aw"
            },
            "views": ["Prometheus Groups", "Prometheus Permissions", "Prometheus Roles", 
                     "Prometheus Permission Set", "Prometheus Set Role Group Ver"]
        }
    ],
    "PROD": [
        {
            "config": {
                "host": "172.30.1.22",
                "port": "5432",
                "dbname": "DB_PROMETHEUS_DOMAIN",
                "user": "prometheus-admin",
                "password": "bos9nJ2We}p3Rba}"
            },
            "views": ["Prometheus Entities", "Prometheus Data Items"]
        },
        {
            "config": {
                "host": "172.30.1.22",
                "port": "5432",
                "dbname": "DB_PROMETHEUS_AUTH",
                "user": "prometheus-admin",
                "password": "bos9nJ2We}p3Rba}"
            },
            "views": ["Prometheus Groups", "Prometheus Permissions", "Prometheus Roles", 
                     "Prometheus Permission Set", "Prometheus Set Role Group Ver"]
        }
    ]
}
