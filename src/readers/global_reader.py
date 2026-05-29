"""
Reads the 'Global Attributes' sheet and returns a dictionary {key: value}.

The sheet is scanned with openpyxl to locate the start/end label rows
(merged cells are handled transparently, since the value lives in the
anchor cell), then only the data rows between those labels are iterated.
"""

from openpyxl.worksheet.worksheet import Worksheet

import config
from src.utils.excel_utils import open_workbook, find_label_rows, cast_value
from src.utils.logger import get_logger

logger = get_logger(__name__)

# Column indices (0-based) within each row tuple returned by iter_rows.
_COL_KEY = 2       # Column C: Key
_COL_VALUE = 3     # Column D: Value
_COL_DATATYPE = 4  # Column E: Data Type


def read_global_attributes() -> dict:
    """
    Opens the CIQ Excel file, locates the Global Attributes table,
    and returns a dictionary mapping each Key to its cast Value.

    Returns:
        dict: {key_string: cast_value, ...}

    Raises:
        FileNotFoundError: If the input Excel file does not exist.
        ValueError: If the expected labels are not found in the sheet.
    """
    input_file = config.INPUT_FILE
    logger.info(f"Reading Global Attributes from: {input_file}")

    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    wb = open_workbook(str(input_file))
    ws: Worksheet = wb[config.SHEET_GLOBAL_ATTRIBUTES]

    start_row, end_row = find_label_rows(
        ws,
        start_label=config.LABEL_GLOBAL_START,
        end_label=config.LABEL_GLOBAL_END,
    )

    # The start label is followed by a header row, so the data begins two rows
    # below it and ends on the row before the end label.
    data_start = start_row + 2
    data_end = end_row - 1

    logger.debug(f"Data rows: {data_start} to {data_end}")

    result: dict = {}
    skipped: int = 0

    for row in ws.iter_rows(
        min_row=data_start,
        max_row=data_end,
        values_only=True,
    ):
        key       = row[_COL_KEY]
        raw_value = row[_COL_VALUE]
        data_type = row[_COL_DATATYPE]

        # Skip rows where Key or Data Type is missing
        if key is None or data_type is None:
            skipped += 1
            logger.debug(f"Skipped incomplete row: {row}")
            continue

        clean_key = str(key).strip()
        if not clean_key:
            skipped += 1
            continue

        result[clean_key] = cast_value(raw_value, str(data_type))

    logger.info(
        f"Global Attributes parsed: {len(result)} entries, {skipped} rows skipped."
    )

    wb.close()
    return {"commonData": result}