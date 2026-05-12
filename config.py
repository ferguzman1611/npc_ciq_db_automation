"""
config.py
---------
Centralizes all configuration values for the CIQ Reader project.
Values are loaded from the .env file via python-dotenv.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ─── Base paths ────────────────────────────────────────────────────────────────
BASE_DIR    = Path(__file__).resolve().parent
INPUT_DIR   = BASE_DIR / "input"
OUTPUT_DIR  = BASE_DIR / "output"
LOGS_DIR    = BASE_DIR / "logs"

# ─── Input file ────────────────────────────────────────────────────────────────
INPUT_FILENAME = os.getenv("INPUT_FILENAME", "TMO_NPC_Voice_CIQDBInputSheet.xlsm")
INPUT_FILE     = INPUT_DIR / INPUT_FILENAME

# ─── Output files ──────────────────────────────────────────────────────────────
OUTPUT_FILENAME_GLOBAL   = os.getenv("OUTPUT_FILENAME_GLOBAL",   "global_attributes.json")
OUTPUT_FILENAME_REGIONAL = os.getenv("OUTPUT_FILENAME_REGIONAL", "regional_attributes.json")
OUTPUT_FILENAME_NETWORK  = os.getenv("OUTPUT_FILENAME_NETWORK",  "network_elements.json")
OUTPUT_FILE_GLOBAL       = OUTPUT_DIR / OUTPUT_FILENAME_GLOBAL
OUTPUT_FILE_REGIONAL     = OUTPUT_DIR / OUTPUT_FILENAME_REGIONAL
OUTPUT_FILE_NETWORK      = OUTPUT_DIR / OUTPUT_FILENAME_NETWORK

# ─── MongoDB ───────────────────────────────────────────────────────────────────
from datetime import datetime as _dt
_host       = os.getenv("MONGODB_HOST", "127.0.0.1")
_port       = os.getenv("MONGODB_PORT", "27017")
_default_db = f"database_{_dt.now().strftime('%Y%m%d_%H%M%S')}"

MONGO_URI                    = f"mongodb://{_host}:{_port}"
MONGO_DB_NAME                = os.getenv("MONGODB_DATABASE", _default_db)
MONGODB_ENABLED              = os.getenv("MONGODB_ENABLED", "false").strip().lower() == "true"
MONGO_COLLECTION_GLOBAL      = os.getenv("MONGO_COLLECTION_GLOBAL",   "global_attributes")
MONGO_COLLECTION_REGIONAL    = os.getenv("MONGO_COLLECTION_REGIONAL", "regional_attributes")
MONGO_COLLECTION_NETWORK     = os.getenv("MONGO_COLLECTION_NETWORK",  "network_elements")

# ─── Excel sheet names ─────────────────────────────────────────────────────────
SHEET_GLOBAL_ATTRIBUTES    = "Global Attributes"
SHEET_REGIONAL_ATTRIBUTES  = "Regional Attributes"

# ─── Labels used to delimit data tables in Excel sheets ────────────────────────
# Each sheet has a start label (just before the header row)
# and an end label (just after the last data row).
LABEL_GLOBAL_START    = "Global attributes definition"
LABEL_GLOBAL_END      = "Continue as needed"   # partial match — avoids issues with special chars (…)
LABEL_REGIONAL_START  = "Regional attributes definition"
LABEL_REGIONAL_END    = "Continue as needed"

# ─── Data type constants ───────────────────────────────────────────────────────
# Only BOOLEAN is cast natively; everything else is coerced to str.
DTYPE_BOOLEAN = "Boolean"
DTYPE_LIST    = "List"