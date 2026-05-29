"""Configuration values for the pipeline, loaded from the .env file."""

import os
from pathlib import Path
from datetime import datetime

from dotenv import load_dotenv

load_dotenv()

# Base paths
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
LOGS_DIR = BASE_DIR / "logs"

# Input file
INPUT_FILENAME = os.getenv("INPUT_FILENAME", "TMO_NPC_Voice_CIQDBInputSheet.xlsm")
INPUT_FILE = INPUT_DIR / INPUT_FILENAME

# Output files. When the corresponding variable is unset, the filename
# falls back to a timestamped name so runs do not overwrite each other.
_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_FILENAME_GLOBAL = os.getenv("OUTPUT_FILENAME_GLOBAL", "").strip() or f"global_attributes_{_timestamp}.json"
OUTPUT_FILENAME_REGIONAL = os.getenv("OUTPUT_FILENAME_REGIONAL", "").strip() or f"regional_attributes_{_timestamp}.json"
OUTPUT_FILENAME_NETWORK = os.getenv("OUTPUT_FILENAME_NETWORK", "").strip() or f"network_elements_{_timestamp}.json"
OUTPUT_FILE_GLOBAL = OUTPUT_DIR / OUTPUT_FILENAME_GLOBAL
OUTPUT_FILE_REGIONAL = OUTPUT_DIR / OUTPUT_FILENAME_REGIONAL
OUTPUT_FILE_NETWORK = OUTPUT_DIR / OUTPUT_FILENAME_NETWORK

# MongoDB
_host = os.getenv("MONGODB_HOST", "127.0.0.1")
_port = os.getenv("MONGODB_PORT", "27017")

MONGO_URI = f"mongodb://{_host}:{_port}"
MONGO_DB_NAME = os.getenv("MONGODB_DATABASE", f"database_{_timestamp}")
MONGODB_ENABLED = os.getenv("MONGODB_ENABLED", "false").strip().lower() == "true"
MONGO_COLLECTION_GLOBAL = os.getenv("MONGO_COLLECTION_GLOBAL", "global_attributes")
MONGO_COLLECTION_REGIONAL = os.getenv("MONGO_COLLECTION_REGIONAL", "regional_attributes")
MONGO_COLLECTION_NETWORK = os.getenv("MONGO_COLLECTION_NETWORK", "network_elements")

# Excel sheet names
SHEET_GLOBAL_ATTRIBUTES = "Global Attributes"
SHEET_REGIONAL_ATTRIBUTES = "Regional Attributes"

# Labels delimiting the data table in each sheet. The end label is matched
# partially to stay resilient against trailing special characters (e.g. "…").
LABEL_GLOBAL_START = "Global attributes definition"
LABEL_GLOBAL_END = "Continue as needed"
LABEL_REGIONAL_START = "Regional attributes definition"
LABEL_REGIONAL_END = "Continue as needed"

# Data type constants. Only Boolean is cast natively; everything else is str.
DTYPE_BOOLEAN = "Boolean"
DTYPE_LIST = "List"
DTYPE_DICT = "Dictionary"