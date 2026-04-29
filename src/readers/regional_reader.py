"""
src/readers/regional_reader.py
-------------------------------
Reads the 'Regional Attributes' sheet from the CIQ Excel file
and returns a list of dictionaries, one per unique Region value.

Each dictionary contains all key-value pairs for that region,
plus a 'region' key identifying it.

Example output:
    [
        {"region": "RegionX", "key1": "value1", ...},
        {"region": "RegionY", "key1": "value99", ...},
    ]

Strategy: same hybrid openpyxl approach as global_reader —
openpyxl finds the label boundaries, then iterates data rows
and groups entries by region.
"""

from collections import defaultdict

import config
from src.utils.excel_utils import open_workbook, find_label_rows, cast_value
from src.utils.logger import get_logger

logger = get_logger(__name__)

# Column indices (0-based) within each row tuple returned by iter_rows
_COL_KEY       = 2   # Column C  →  Key
_COL_VALUE     = 3   # Column D  →  Value
_COL_DATATYPE  = 4   # Column E  →  Data Type
_COL_REGION    = 5   # Column F  →  Region


def read_regional_attributes() -> list[dict]:
    """
    Opens the CIQ Excel file, locates the Regional Attributes table,
    and returns a list of dictionaries grouped by Region.

    Returns:
        list[dict]: One dictionary per unique region, each containing
                    all key-value pairs for that region plus a 'region' key.

    Raises:
        FileNotFoundError: If the input Excel file does not exist.
        ValueError: If the expected labels are not found in the sheet.
    """
    input_file = config.INPUT_FILE
    logger.info(f"Reading Regional Attributes from: {input_file}")

    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    wb = open_workbook(str(input_file))
    ws = wb[config.SHEET_REGIONAL_ATTRIBUTES]

    # ── Step 1: locate table boundaries ───────────────────────────────────────
    start_row, end_row = find_label_rows(
        ws,
        start_label=config.LABEL_REGIONAL_START,
        end_label=config.LABEL_REGIONAL_END,
    )

    data_start = start_row + 2   # skip label row + header row
    data_end   = end_row - 1     # stop before "Continue as needed…"

    logger.debug(f"Data rows: {data_start} to {data_end}")

    # ── Step 2: iterate rows and group by region ───────────────────────────────
    # Using defaultdict to accumulate key-value pairs per region
    grouped: dict[str, dict] = defaultdict(dict)
    skipped = 0

    for row in ws.iter_rows(
        min_row=data_start,
        max_row=data_end,
        values_only=True,
    ):
        key       = row[_COL_KEY]
        raw_value = row[_COL_VALUE]
        data_type = row[_COL_DATATYPE]
        region    = row[_COL_REGION]

        # Skip rows with missing required fields
        if any(v is None for v in (key, data_type, region)):
            skipped += 1
            logger.debug(f"Skipped incomplete row: {row}")
            continue

        clean_key    = str(key).strip()
        clean_region = str(region).strip()

        if not clean_key or not clean_region:
            skipped += 1
            continue

        grouped[clean_region][clean_key] = cast_value(raw_value, str(data_type))

    # ── Step 3: build final list, injecting 'region' key into each dict ────────
    result = []
    for region_name, attributes in grouped.items():
        entry = {"region": region_name, **attributes}
        result.append(entry)

    wb.close()

    logger.info(
        f"Regional Attributes parsed: {len(result)} regions, {skipped} rows skipped."
    )

    return result