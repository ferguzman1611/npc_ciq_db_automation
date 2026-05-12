"""
src/utils/excel_utils.py
------------------------
Helper functions for reading and parsing Excel files with openpyxl.
Handles merged cells, label searching, and data type casting.
"""

from typing import Any, Optional, Tuple
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import json
from config import DTYPE_BOOLEAN, DTYPE_LIST
from src.utils.logger import get_logger

logger = get_logger(__name__)


def open_workbook(filepath: str):
    """
    Opens an Excel workbook in read-only mode.
    Supports both .xlsx and .xlsm files.

    Args:
        filepath: Absolute or relative path to the Excel file.

    Returns:
        An openpyxl Workbook object.

    Raises:
        FileNotFoundError: If the file does not exist.
        Exception: If openpyxl cannot open the file.
    """
    logger.debug(f"Opening workbook: {filepath}")
    return load_workbook(filepath, read_only=True, keep_vba=True)


def find_label_rows(
    ws: Worksheet,
    start_label: str,
    end_label: str,
) -> Tuple[int, int]:
    """
    Scans a worksheet row by row to find the rows containing
    the start and end labels that delimit a data table.

    Uses partial matching (case-insensitive) to be resilient
    against special characters (e.g. ellipsis '…' vs '...').

    Args:
        ws:          The openpyxl worksheet to scan.
        start_label: Substring expected in the start-delimiter cell.
        end_label:   Substring expected in the end-delimiter cell.

    Returns:
        A tuple (start_row, end_row) with 1-based row indices.

    Raises:
        ValueError: If either label is not found in the worksheet.
    """
    start_row: Optional[int] = None
    end_row:   Optional[int] = None

    for row in ws.iter_rows(values_only=True):
        for i, cell_value in enumerate(row):
            if cell_value is None:
                continue
            cell_str = str(cell_value).strip()

            if start_row is None and start_label.lower() in cell_str.lower():
                # row is a tuple; openpyxl rows are 1-indexed
                # We compute the row number from the worksheet iteration
                start_row = _get_row_number(ws, cell_str, start_label)
                logger.debug(f"Start label '{start_label}' found at row {start_row}")

            elif start_row is not None and end_label.lower() in cell_str.lower():
                end_row = _get_row_number(ws, cell_str, end_label)
                logger.debug(f"End label '{end_label}' found at row {end_row}")
                break

        if start_row and end_row:
            break

    if start_row is None:
        raise ValueError(f"Start label '{start_label}' not found in worksheet '{ws.title}'.")
    if end_row is None:
        raise ValueError(f"End label '{end_label}' not found in worksheet '{ws.title}'.")

    return start_row, end_row


def _get_row_number(ws: Worksheet, cell_str: str, label: str) -> int:
    """
    Secondary scan to get the actual 1-based row index of a label.
    Called internally by find_label_rows after a match is confirmed.
    """
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and label.lower() in str(cell.value).strip().lower():
                return cell.row
    raise ValueError(f"Label '{label}' not found during row-number lookup.")


def cast_value(value: Any, data_type: str) -> Any:
    """
    Casts a cell value according to the Data Type column rules:

    - Boolean  → Python bool  (True/False)
    - Anything else (String, Integer, Float, etc.) → str

    Leading/trailing whitespace is always stripped.

    Args:
        value:     Raw cell value from openpyxl.
        data_type: String from the 'Data Type' column (e.g. 'String', 'Boolean').

    Returns:
        The value cast to the appropriate Python type.
    """
    if value is None:
        return None

    cleaned = str(value).strip()

    if data_type.strip() == DTYPE_BOOLEAN:
        return cleaned.lower() in ("true", "1", "yes")

    # Try parse as list if value starts with '['
    if data_type.strip() == DTYPE_LIST:
        return _parse_informal_list(cleaned)

    return cleaned


def cast_bool_aware(value: Any) -> Any:
    """
    Casts a value to bool if it looks like a boolean, regardless of Data Type.
    Used for values inside Dictionary and List fields, where the Data Type column
    only describes the container (Dictionary/List), not the inner values.

    Recognizes all variants case-insensitively and with or without outer quotes:
        true, false, "true", "false", 'True', 'False', etc.

    Args:
        value: Raw string value extracted from an Excel cell.

    Returns:
        True/False if the value is a boolean variant, otherwise the original string.
    """
    if value is None:
        return None

    # Strip outer whitespace and surrounding quote characters to normalize
    cleaned = str(value).strip().strip('"').strip("'").strip()

    if cleaned.lower() == "true":
        return True
    if cleaned.lower() == "false":
        return False

    # Not a boolean — return original value stripped of outer whitespace only
    return str(value).strip()

def _parse_informal_list(raw: str) -> list:
    """
    Parses an informal list string like:
        [a, b, "c,d", e]
        [a,\n b,\n c]

    Rules:
    - Strips the outer brackets
    - Removes newlines and extra whitespace
    - Respects quoted elements (e.g. "BACKUP,TASKS" → one element)
    - Returns a list of stripped strings
    """
    import csv
    import io

    # Remove outer brackets and normalize newlines/whitespace
    inner = raw.strip()[1:-1]
    inner = " ".join(inner.splitlines())  # colapsa saltos de línea

    # Use csv.reader to respect quoted substrings
    reader = csv.reader(io.StringIO(inner), skipinitialspace=True)
    elements = next(reader, [])

    return [el.strip() for el in elements if el.strip()]