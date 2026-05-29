"""
Reads all '*-InputTable' sheets and returns a flat list of dictionaries,
one per network element (row).

Sheet structure (fixed rows 1-5):
    Row 1: Quantity   - sub-elements per attribute. 1 for scalars, N value
                        columns for List, N key/value pairs (N*2 columns)
                        for Dictionary.
    Row 2: Data type  - String, Boolean, Integer, Float, Dictionary or List.
    Row 3: Provided by - ignored.
    Row 4: Key        - attribute name. Column A holds the row sequence
                        number and is always discarded.
    Row 5: Sub-headers - Value1 for scalars; Key1/Value1/Key2/Value2... for
                        Dictionary; Value1/Value2/... for List.
    Row 6+: one network element per row.

Data type handling:
    Scalar      cast to a stripped str.
    Boolean     cast to a Python bool.
    Dictionary  built from Key/Value sub-column pairs in the same row; a pair
                is discarded when either cell is None; inner values go through
                cast_bool_aware().
    List        each Value sub-column is parsed as JSON; unparseable cells
                become "INVALID_JSON".
"""

import json
from typing import Any

from openpyxl import load_workbook

import config
from src.utils.excel_utils import cast_value, cast_bool_aware
from src.utils.logger import get_logger

logger = get_logger(__name__)

# Fixed row indices (0-based into the all_rows list).
_ROW_QUANTITY = 0   # Row 1
_ROW_DATATYPE = 1   # Row 2
_ROW_KEY = 3        # Row 4
_ROW_SUBHEADER = 4  # Row 5
_DATA_START_IDX = 5  # Row 6

_DTYPE_DICT = "Dictionary"
_DTYPE_LIST = "List"
_QUOTED_EMPTY = '""'
_SKIP_COL = 0  # Column A: row sequence number, always discarded.


def read_network_elements() -> list[dict]:
    """
    Reads every '*-InputTable' sheet in the workbook and returns a single flat
    list of network element documents (the per-sheet results are concatenated).

    Raises:
        FileNotFoundError: If the configured input file does not exist.
    """
    input_file = config.INPUT_FILE
    logger.info(f"Reading Network Elements from: {input_file}")

    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    # read_only streams rows without loading the whole workbook into memory;
    # data_only returns computed values rather than formulas; keep_vba is needed
    # because the source file is a macro-enabled .xlsm.
    wb = load_workbook(str(input_file), read_only=True, keep_vba=True, data_only=True)

    # Network element tables live on sheets whose name ends in "-InputTable".
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


def _parse_sheet(ws, sheet_name: str) -> list[dict]:
    """Parses a single InputTable sheet into a list of network element dicts."""
    all_rows = list(ws.iter_rows(values_only=True))

    # A valid sheet needs at least the fixed header rows plus one data row.
    if len(all_rows) < _DATA_START_IDX + 1:
        logger.warning(f"Sheet '{sheet_name}' has too few rows; skipped.")
        return []

    # The first rows describe the table layout (see the module docstring); the
    # remaining rows are the actual network elements.
    quantity_row = all_rows[_ROW_QUANTITY]
    datatype_row = all_rows[_ROW_DATATYPE]
    key_row = all_rows[_ROW_KEY]
    subheader_row = all_rows[_ROW_SUBHEADER]
    data_rows = all_rows[_DATA_START_IDX:]

    # Derive the column layout once, then apply it to every data row.
    schema = _build_schema(quantity_row, datatype_row, key_row, subheader_row)

    elements = []
    for row_tuple in data_rows:
        # A fully empty row marks the end of the table; trailing rows below it
        # (if any) are ignored.
        if all(v is None for v in row_tuple):
            break
        doc = _parse_row(row_tuple, schema)
        if doc:
            elements.append(doc)

    return elements


def _build_schema(
    quantity_row: tuple,
    datatype_row: tuple,
    key_row: tuple,
    subheader_row: tuple,
) -> list[dict]:
    """
    Builds a column schema by scanning left to right.

    Column consumption per attribute: a scalar takes 1 column, a List takes
    its quantity N of columns, and a Dictionary takes N*2 columns (N key/value
    pairs).

    Returns:
        A list of attribute descriptors, each a dict with keys: 'name', 'dtype',
        'col' (the attribute's first column index), 'quantity', and 'sub_cols'
        (the list of column indices the attribute spans).
    """
    schema = []
    col = 0
    col_count = len(key_row)

    # Walk the header columns left to right. Each attribute advances the cursor
    # by however many columns it occupies, so multi-column List/Dictionary
    # attributes are consumed in a single step.
    while col < col_count:
        # Column A is the row sequence number, not an attribute.
        if col == _SKIP_COL:
            col += 1
            continue

        attr_name = key_row[col]
        dtype = datatype_row[col]
        quantity = quantity_row[col]

        # Columns without a name or data type are spacers; skip them one by one.
        if attr_name is None or dtype is None:
            col += 1
            continue

        attr_name = str(attr_name).strip()
        dtype = str(dtype).strip()

        if not attr_name:
            col += 1
            continue

        # Quantity defaults to 1 when missing or non-numeric (i.e. a scalar).
        try:
            quantity = int(quantity) if quantity is not None else 1
        except (ValueError, TypeError):
            quantity = 1

        # A Dictionary stores N key/value pairs across 2N columns; every other
        # type occupies exactly `quantity` columns. min() guards against a
        # declared span that would run past the end of the row.
        col_span = quantity * 2 if dtype == _DTYPE_DICT else quantity
        sub_cols = list(range(col, min(col + col_span, col_count)))

        schema.append({
            'name': attr_name,
            'dtype': dtype,
            'col': col,
            'quantity': quantity,
            'sub_cols': sub_cols,
        })

        col += col_span

    return schema


def _parse_row(row: tuple, schema: list[dict]) -> dict:
    """
    Converts one data row into a network element dict using the schema.

    Attributes whose value is missing (None) are omitted from the result rather
    than stored as null, so documents only contain the fields actually present.
    """
    doc = {}

    for attr in schema:
        name = attr['name']
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
            # Scalar: a single cell at the attribute's first column.
            raw = _safe_get(row, attr['col'])
            cleaned = _clean_scalar(raw)
            if cleaned is None:
                continue
            doc[name] = cast_value(cleaned, dtype)

    return doc


def _extract_dict(row: tuple, sub_cols: list[int]) -> dict | None:
    """
    Builds a dict from alternating Key/Value sub-columns laid out as
    [key1_col, val1_col, key2_col, val2_col, ...].

    A pair is discarded when its key or value cell is None; a quoted empty
    value is kept as an empty string; inner values go through cast_bool_aware().
    """
    result = {}

    # sub_cols alternates key/value columns, so pair them up two at a time.
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


def _extract_list(row: tuple, sub_cols: list[int]) -> list | None:
    """
    Builds a list from Value sub-columns. Each non-None cell is parsed as
    JSON; unparseable cells become "INVALID_JSON".
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


def _safe_get(row: tuple, col: int) -> Any:
    """Returns row[col], or None when the index is out of bounds."""
    return row[col] if col < len(row) else None


def _clean_scalar(value: Any) -> str | None:
    """
    Normalizes a raw scalar cell value: None and empty cells become None (the
    key is then excluded from the document), a quoted empty string becomes an
    empty string, and anything else is returned stripped.
    """
    if value is None:
        return None
    s = str(value).strip()
    if s == _QUOTED_EMPTY:
        return ""
    return s if s else None


def _safe_parse_json(text: str) -> Any:
    """Parses a string as JSON, returning "INVALID_JSON" if it cannot be parsed."""
    try:
        return json.loads(text)
    except (json.JSONDecodeError, ValueError):
        logger.warning(f"Could not parse JSON: {text!r}; stored as 'INVALID_JSON'")
        return "INVALID_JSON"