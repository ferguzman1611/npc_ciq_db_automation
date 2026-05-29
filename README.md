# npc_ciq_db_automation

Python pipeline that reads a CIQ Excel file (`.xlsm`) and produces structured JSON files and MongoDB documents for three data domains: Global Attributes, Regional Attributes, and Network Elements.

---

## Project structure

```
npc_ciq_db_automation/
├── input/                              # Place the input Excel file here (git-ignored)
├── output/                             # Generated JSON files (git-ignored)
├── logs/                               # Timestamped log files (git-ignored)
├── src/
│   ├── readers/
│   │   ├── global_reader.py            # Reads the "Global Attributes" sheet
│   │   ├── regional_reader.py          # Reads the "Regional Attributes" sheet
│   │   └── network_elements_reader.py  # Reads all "*-InputTable" sheets
│   ├── writers/
│   │   ├── json_writer.py              # Writes output JSON files
│   │   └── mongo_writer.py             # Writes to MongoDB collections
│   └── utils/
│       ├── excel_utils.py              # Shared Excel helpers (label search, casting)
│       └── logger.py                   # Logger configuration
├── config.py                           # Loads .env and exposes all constants
├── main.py                             # Entry point that orchestrates the pipelines
├── .env                                # Local environment config (git-ignored)
├── .env.example                        # Template for .env
├── .gitignore
├── requirements.txt
└── README.md
```

---

## Requirements

- Python 3.10+
- MongoDB instance accessible from the host machine

---

## Setup

### 1. Clone the repository

```bash
git clone <repo-url>
cd npc_ciq_db_automation
```

### 2. Create and activate a virtual environment

```bash
python3 -m venv .venv
source .venv/bin/activate      # Linux / macOS
.venv\Scripts\activate         # Windows
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure environment variables

Copy `.env.example` to `.env` and fill in your values:

```bash
cp .env.example .env
```

Edit `.env` with your MongoDB connection details and desired file names. See the [Environment variable reference](#environment-variable-reference) section below for all available options.

> If `MONGODB_DATABASE` is not set, the database name defaults to `database_<YYYYMMDD_HHMMSS>` (timestamp of the run).

### 5. Place the input file

Copy the CIQ Excel file into the `input/` folder:

```bash
cp /path/to/TMO_NPC_Voice_CIQDBInputSheet.xlsm input/
```

---

## Usage

Run all three pipelines at once:

```bash
python main.py
```

Or run individual pipelines using flags:

```bash
python main.py --global      # Global Attributes only
python main.py --regional    # Regional Attributes only
python main.py --network     # Network Elements only
```

Flags can also be combined:

```bash
python main.py --global --network
```

---

## Outputs

Each pipeline produces two outputs:

| Pipeline | JSON file | MongoDB collection |
|---|---|---|
| Global Attributes | `output/global_attributes_<timestamp>.json` | `global_attributes` |
| Regional Attributes | `output/regional_attributes_<timestamp>.json` | `regional_attributes` |
| Network Elements | `output/network_elements_<timestamp>.json` | `network_elements` |

By default, output filenames include a `YYYYMMDD_HHMMSS` timestamp so runs do not overwrite each other. Set `OUTPUT_FILENAME_*` in `.env` to use fixed names instead. Collection names are also configurable via `.env`.

---

## MongoDB write strategy

| Pipeline | Strategy |
|---|---|
| Global Attributes | Single document, replaced on each run (upsert) |
| Regional Attributes | Full collection refresh: dropped and reinserted on each run |
| Network Elements | Upsert per document by `Node_Name`: existing documents updated, new ones inserted |

---

## Excel sheet conventions

### Global Attributes / Regional Attributes

- Label `"Global/Regional attributes definition"` marks the start of the table.
- Label `"Continue as needed…"` marks the end.
- Columns used: `Key`, `Value`, `Data Type` (and `Region` for Regional).

### `*-InputTable` sheets (Network Elements)

| Row | Content |
|---|---|
| 1 | `Quantity`: number of sub-elements per attribute (pairs for Dictionary, columns for List) |
| 2 | `Data type`: `String`, `Boolean`, `Integer`, `Float`, `Dictionary`, or `List` |
| 4 | `Key`: attribute names. Column A (row sequence number) is always discarded |
| 5 | Sub-headers: `Value1` for scalars; `Key1/Value1/Key2/Value2…` for Dictionary; `Value1/Value2…` for List |
| 6+ | One network element per row |

### Data type rules

- `Boolean` becomes a Python `bool`. Recognized case-insensitively, with or without surrounding quotes: `true`, `"true"`, `True`, `false`, etc.
- Inner values of `Dictionary` / `List` are also cast to `bool` when they match the above.
- All other types become `str`.
- `List` elements: each cell is parsed as a JSON string. Unparseable values are stored as `"INVALID_JSON"`.
- A `None` cell causes the key to be excluded from the document entirely.
- A quoted empty string in Excel (`""`) is stored as an empty string `""`.

---

## Logs

Each run creates a timestamped log file in `logs/`:

```
logs/2026-04-28_12-46-42.log
```

`INFO` level is shown in the console. `DEBUG` level is written to the log file only.

---

## Environment variable reference

| Variable | Default | Description |
|---|---|---|
| `MONGODB_HOST` | `127.0.0.1` | MongoDB host |
| `MONGODB_PORT` | `27017` | MongoDB port |
| `MONGODB_DATABASE` | `database_<timestamp>` | Target database name |
| `MONGO_COLLECTION_GLOBAL` | `global_attributes` | Collection for Global Attributes |
| `MONGO_COLLECTION_REGIONAL` | `regional_attributes` | Collection for Regional Attributes |
| `MONGO_COLLECTION_NETWORK` | `network_elements` | Collection for Network Elements |
| `MONGODB_ENABLED` | `false` | Set to `true` to write to MongoDB; otherwise only JSON files are produced |
| `INPUT_FILENAME` | `TMO_NPC_Voice_CIQDBInputSheet.xlsm` | Input Excel filename (must be in `input/`) |
| `OUTPUT_FILENAME_GLOBAL` | `global_attributes_<timestamp>.json` | Output filename for Global Attributes |
| `OUTPUT_FILENAME_REGIONAL` | `regional_attributes_<timestamp>.json` | Output filename for Regional Attributes |
| `OUTPUT_FILENAME_NETWORK` | `network_elements_<timestamp>.json` | Output filename for Network Elements |