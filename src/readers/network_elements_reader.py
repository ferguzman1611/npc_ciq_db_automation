"""
src/readers/network_elements_reader.py
---------------------------------------
Reads all '*-InputTable' sheets from the CIQ Excel file and returns
a flat list of dictionaries, one per network element (row).

Sheet structure (fixed rows 1-5):
    Row 1 → Quantity:   how many sub-elements each attribute has.
                        For scalars: always 1 (= 1 column).
                        For List:    N = number of Value columns.
                        For Dict:    N = number of Key/Value PAIRS
                                     → occupies N*2 columns.
    Row 2 → Data type:  attribute type (String, Boolean, Dictionary, List, ...)
    Row 3 → Provided by (ignored)
    Row 4 → Key:        attribute name. Column A is the row sequence
                        number — always discarded.
    Row 5 → N°\nValue:  sub-headers (Value1 for scalars;
                        Key1/Value1/Key2/Value2... for dict;
                        Value1/Value2/... for list)
    Row 6+ → one network element per row (horizontal reading)

Data type handling:
    Scalar (String, Integer, Float, etc.) → str  (stripped)
    Boolean  → Python bool
    Dictionary → built from Key/Value sub-column pairs in the same row.
                 Quantity N → N pairs → N*2 columns.
                 Pair discarded if key OR value cell is None.
                 Inner values cast with cast_bool_aware().
    List     → each Value sub-column parsed as JSON string.
                 Quantity N → N columns.
                 json.loads() handles inner booleans natively.
                 Unparseable cells → "INVALID_JSON".
"""

import json
from typing import Any

from openpyxl import load_workbook

import config
from src.utils.excel_utils import cast_value, cast_bool_aware
from src.utils.logger import get_logger

logger = get_logger(__name__)

# ── Fixed row indices (0-based into all_rows list) ─────────────────────────────
_ROW_QUANTITY   = 0   # Row 1
_ROW_DATATYPE   = 1   # Row 2
_ROW_KEY        = 3   # Row 4
_ROW_SUBHEADER  = 4   # Row 5
_DATA_START_IDX = 5   # Row 6

_DTYPE_DICT   = "Dictionary"
_DTYPE_LIST   = "List"
_QUOTED_EMPTY = '""'
_SKIP_COL     = 0    # Column A: row sequence number, always discarded


# ── Public entry point ────────────────────────────────────────────────────────

def read_network_elements() -> list[dict]:
    """
    Reads every '*-InputTable' sheet and returns a flat list of
    network element documents.
    """
    input_file = config.INPUT_FILE
    logger.info(f"Reading Network Elements from: {input_file}")

    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    wb = load_workbook(str(input_file), read_only=True, keep_vba=True, data_only=True)

    input_sheets = [s for s in wb.sheetnames if s.endswith("-InputTable")]
    logger.info(f"Found {len(input_sheets)} InputTable sheet(s): {input_sheets}")

    all_elements: list[dict] = []

    for sheet_name in input_sheets:
        ws = wb[sheet_name]
        elements = _parse_sheet(ws, sheet_name)
        logger.info(f"  {sheet_name}: {len(elements)} network element(s) parsed.")
        all_elements.extend(elements)

    wb.close()
    logger.info(f"Total network elements parsed: {len(all_elements)}")
    return all_elements


# ── Sheet parser ──────────────────────────────────────────────────────────────

def _parse_sheet(ws, sheet_name: str) -> list[dict]:
    """Parses a single InputTable sheet into a list of network element dicts."""
    all_rows = list(ws.iter_rows(values_only=True))

    if len(all_rows) < _DATA_START_IDX + 1:
        logger.warning(f"Sheet '{sheet_name}' has too few rows — skipped.")
        return []

    quantity_row  = all_rows[_ROW_QUANTITY]
    datatype_row  = all_rows[_ROW_DATATYPE]
    key_row       = all_rows[_ROW_KEY]
    subheader_row = all_rows[_ROW_SUBHEADER]
    data_rows     = all_rows[_DATA_START_IDX:]

    schema = _build_schema(quantity_row, datatype_row, key_row, subheader_row)

    elements = []
    for row_tuple in data_rows:
        if all(v is None for v in row_tuple):
            break
        doc = _parse_row(row_tuple, schema)
        if doc:
            elements.append(doc)

    return elements


# ── Schema builder ─────────────────────────────────────────────────────────────

def _build_schema(
    quantity_row: tuple,
    datatype_row: tuple,
    key_row: tuple,
    subheader_row: tuple,
) -> list[dict]:
    """
    Builds a column schema scanning left to right.

    Key rule for column consumption:
        Scalar     → quantity = 1  → consumes 1 column
        List       → quantity = N  → consumes N columns
        Dictionary → quantity = N  → consumes N*2 columns (N Key/Value pairs)
    """
    schema = []
    col = 0
    col_count = len(key_row)

    while col < col_count:
        if col == _SKIP_COL:
            col += 1
            continue

        attr_name = key_row[col]
        dtype     = datatype_row[col]
        quantity  = quantity_row[col]

        if attr_name is None or dtype is None:
            col += 1
            continue

        attr_name = str(attr_name).strip()
        dtype     = str(dtype).strip()

        if not attr_name:
            col += 1
            continue

        try:
            quantity = int(quantity) if quantity is not None else 1
        except (ValueError, TypeError):
            quantity = 1

        # ── Calculate how many columns this attribute occupies ─────────────────
        if dtype == _DTYPE_DICT:
            # N pairs → N*2 columns (Key1, Val1, Key2, Val2, ...)
            col_span = quantity * 2
        else:
            # Scalar: 1 col. List: N cols.
            col_span = quantity

        sub_cols = list(range(col, min(col + col_span, col_count)))

        schema.append({
            'name':     attr_name,
            'dtype':    dtype,
            'col':      col,
            'quantity': quantity,
            'sub_cols': sub_cols,
        })

        col += col_span

    return schema


# ── Row parser ─────────────────────────────────────────────────────────────────

def _parse_row(row: tuple, schema: list[dict]) -> dict:
    """Converts one data row into a network element dict using the schema."""
    doc = {}

    for attr in schema:
        name  = attr['name']
        dtype = attr['dtype']

        if dtype == _DTYPE_DICT:
            value = _extract_dict(row, attr['sub_cols'])
            if value is not None:
                doc[name] = value

        elif dtype == _DTYPE_LIST:
            value = _extract_list(row, attr['sub_cols'])
            if value is not None:
                doc[name] = value

        else:
            raw = _safe_get(row, attr['col'])
            cleaned = _clean_scalar(raw)
            if cleaned is None:
                continue
            doc[name] = cast_value(cleaned, dtype)

    return doc


# ── Dictionary extractor ───────────────────────────────────────────────────────

def _extract_dict(row: tuple, sub_cols: list[int]) -> dict | None:
    """
    Builds a dict from alternating Key/Value sub-columns.
    sub_cols layout: [key1_col, val1_col, key2_col, val2_col, ...]
    Rules:
      - key cell None  → discard pair
      - value cell None → discard pair
      - value cell ""  → include with empty string
      - inner values cast with cast_bool_aware()
    """
    result = {}

    pairs = [(sub_cols[i], sub_cols[i + 1])
             for i in range(0, len(sub_cols) - 1, 2)]

    for key_col, val_col in pairs:
        raw_key = _safe_get(row, key_col)
        raw_val = _safe_get(row, val_col)

        if raw_key is None or raw_val is None:
            continue

        clean_key = _clean_scalar(raw_key)
        if clean_key is None:
            continue

        clean_val = _clean_scalar(raw_val)
        if clean_val is None:
            continue

        result[clean_key] = cast_bool_aware(clean_val)

    return result if result else None


# ── List extractor ─────────────────────────────────────────────────────────────

def _extract_list(row: tuple, sub_cols: list[int]) -> list | None:
    """
    Builds a list from Value sub-columns.
    Each non-None cell is parsed as JSON.
    Unparseable cells → "INVALID_JSON".
    """
    result = []

    for col in sub_cols:
        raw = _safe_get(row, col)
        if raw is None:
            continue
        cleaned = str(raw).strip()
        if not cleaned:
            continue
        result.append(_safe_parse_json(cleaned))

    return result if result else None


# ── Helpers ────────────────────────────────────────────────────────────────────

def _safe_get(row: tuple, col: int) -> Any:
    """Returns row[col] or None if out of bounds."""
    return row[col] if col < len(row) else None


def _clean_scalar(value: Any) -> str | None:
    """
    Normalizes a raw scalar cell value:
      - None  → None  (key excluded from doc)
      - '""'  → ""    (quoted empty → empty string)
      - other → stripped string, None if empty after strip
    """
    if value is None:
        return None
    s = str(value).strip()
    if s == _QUOTED_EMPTY:
        return ""
    return s if s else None


def _safe_parse_json(text: str) -> Any:
    """
    Parses a string as JSON. Returns "INVALID_JSON" if unparseable.
    json.loads() natively handles true/false → Python bool.
    """
    try:
        return json.loads(text)
    except (json.JSONDecodeError, ValueError):
        logger.warning(f"Could not parse JSON: {text!r} → 'INVALID_JSON'")
        return "INVALID_JSON"