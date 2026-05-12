"""
src/utils/excel_utils.py
------------------------
Helper functions for reading and parsing Excel files with openpyxl.
Handles merged cells, label searching, and data type casting.
"""

import csv
import io
from typing import Any, Optional, Tuple
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from config import DTYPE_BOOLEAN, DTYPE_LIST, DTYPE_DICT
from src.utils.logger import get_logger

logger = get_logger(__name__)


def open_workbook(filepath: str):
    """
    Opens an Excel workbook in read-only mode.
    Supports both .xlsx and .xlsm files.
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
    """
    start_row: Optional[int] = None
    end_row:   Optional[int] = None

    for row in ws.iter_rows(values_only=True):
        for i, cell_value in enumerate(row):
            if cell_value is None:
                continue
            cell_str = str(cell_value).strip()

            if start_row is None and start_label.lower() in cell_str.lower():
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

    - Boolean    → Python bool
    - List       → Python list  (parsed via _parse_informal_list)
    - Dictionary → Python dict  (parsed via _parse_informal_dict)
    - Anything else → str

    Numbers inside List/Dictionary are always cast to str.
    Booleans inside List/Dictionary are cast to Python bool.
    """
    if value is None:
        return None

    cleaned = str(value).strip()
    dtype   = data_type.strip()

    if dtype == DTYPE_BOOLEAN:
        return cleaned.lower() in ("true", "1", "yes")

    if dtype == DTYPE_LIST:
        return _parse_informal_list(cleaned)

    if dtype == DTYPE_DICT:
        result = _parse_informal_dict(cleaned)
        if result is None:
            logger.warning(
                f"Could not parse Dictionary value — stored as string: {cleaned!r}"
            )
            return cleaned
        return result

    return cleaned


def cast_bool_aware(value: Any) -> Any:
    """
    Casts a value to bool if it looks like a boolean, regardless of Data Type.
    Used for values inside Dictionary and List fields.

    Recognizes all variants case-insensitively and with or without outer quotes:
        true, false, "true", "false", 'True', 'False', etc.
    """
    if value is None:
        return None

    cleaned = str(value).strip().strip('"').strip("'").strip()

    if cleaned.lower() == "true":
        return True
    if cleaned.lower() == "false":
        return False

    return str(value).strip()


# ── List parser ────────────────────────────────────────────────────────────────

def _parse_informal_list(raw: str) -> list:
    """
    Parses an informal list string into a Python list.

    Handles:
      - Optional outer brackets [ ]
      - Newlines and extra whitespace
      - Quoted elements that contain commas: "a,b" → one element
      - Each element cast with _cast_scalar (bool-aware, numbers → str)

    Examples:
      [a, b, "c,d"]          → ["a", "b", "c,d"]
      [a,\\n b,\\n c]         → ["a", "b", "c"]
      [1, 2, true]           → ["1", "2", True]
    """
    # Strip outer brackets if present
    stripped = raw.strip()
    if stripped.startswith("[") and stripped.endswith("]"):
        stripped = stripped[1:-1]

    # Collapse newlines and normalize whitespace
    normalized = " ".join(stripped.splitlines())

    # Use csv.reader to respect quoted substrings containing commas
    reader   = csv.reader(io.StringIO(normalized), skipinitialspace=True)
    elements = next(reader, [])

    return [_cast_scalar(el.strip()) for el in elements if el.strip()]


# ── Dict parser ────────────────────────────────────────────────────────────────

def _parse_informal_dict(raw: str) -> dict | None:
    """
    Parses an informal dict/JSON string into a Python dict.

    Handles:
      - Standard JSON (with or without quotes on keys/values)
      - Newlines and extra whitespace
      - Unquoted string values (including those with spaces):
            "key": some value with spaces  → {"key": "some value with spaces"}
      - Unquoted keys:
            key: "value"                   → {"key": "value"}
      - Nested dicts and lists (recursively parsed)
      - Numbers → always str
      - Booleans → Python bool
      - Missing commas or structural typos → returns None (logged as WARNING)

    Returns:
      Parsed dict, or None if the structure is invalid.
    """
    # Collapse newlines and normalize whitespace
    normalized = " ".join(raw.strip().splitlines()).strip()

    # Must be wrapped in { }
    if not (normalized.startswith("{") and normalized.endswith("}")):
        return None

    try:
        tokens = _tokenize(normalized)
        result, pos = _parse_dict_tokens(tokens, 0)
        return result
    except Exception as exc:
        logger.warning(f"Dict parse error ({exc}): {raw!r}")
        return None


# ── Tokenizer ──────────────────────────────────────────────────────────────────

def _tokenize(text: str) -> list[str]:
    """
    Splits the input into tokens:
      - Structural chars: { } [ ] , :
      - Quoted strings (preserving content between double quotes)
      - Unquoted words (anything else, split on structural chars)

    Whitespace between tokens is discarded.
    """
    tokens = []
    i = 0
    n = len(text)

    while i < n:
        ch = text[i]

        if ch in ' \t\r\n':
            i += 1
            continue

        if ch in '{}[],: ':
            tokens.append(ch)
            i += 1
            continue

        if ch == '"':
            # Quoted string — read until closing unescaped "
            j = i + 1
            while j < n:
                if text[j] == '\\':
                    j += 2
                    continue
                if text[j] == '"':
                    break
                j += 1
            tokens.append(text[i:j + 1])   # includes surrounding quotes
            i = j + 1
            continue

        # Unquoted token — read until structural char
        j = i
        while j < n and text[j] not in '{}[],:"\t\r\n':
            j += 1
        word = text[i:j].strip()
        if word:
            tokens.append(word)
        i = j

    return tokens


# ── Recursive token parsers ────────────────────────────────────────────────────

def _parse_dict_tokens(tokens: list[str], pos: int) -> tuple[dict, int]:
    """
    Parses a dict starting at tokens[pos] (which must be '{').
    Returns (parsed_dict, next_pos).
    """
    assert tokens[pos] == '{', f"Expected '{{' at pos {pos}, got {tokens[pos]!r}"
    pos += 1
    result = {}

    while pos < len(tokens) and tokens[pos] != '}':
        # ── Parse key ──────────────────────────────────────────────────────────
        key_token = tokens[pos]
        pos += 1
        key = _unquote(key_token)

        # Expect ':'
        if pos >= len(tokens) or tokens[pos] != ':':
            raise ValueError(f"Expected ':' after key {key!r}, got {tokens[pos] if pos < len(tokens) else 'EOF'!r}")
        pos += 1

        # ── Parse value ────────────────────────────────────────────────────────
        value, pos = _parse_value_tokens(tokens, pos)
        result[key] = value

        # Optional comma
        if pos < len(tokens) and tokens[pos] == ',':
            pos += 1

    # Consume closing '}'
    if pos < len(tokens) and tokens[pos] == '}':
        pos += 1

    return result, pos


def _parse_list_tokens(tokens: list[str], pos: int) -> tuple[list, int]:
    """
    Parses a list starting at tokens[pos] (which must be '[').
    Returns (parsed_list, next_pos).
    """
    assert tokens[pos] == '[', f"Expected '[' at pos {pos}, got {tokens[pos]!r}"
    pos += 1
    result = []

    while pos < len(tokens) and tokens[pos] != ']':
        value, pos = _parse_value_tokens(tokens, pos)
        result.append(value)

        # Optional comma
        if pos < len(tokens) and tokens[pos] == ',':
            pos += 1

    # Consume closing ']'
    if pos < len(tokens) and tokens[pos] == ']':
        pos += 1

    return result, pos


def _parse_value_tokens(tokens: list[str], pos: int) -> tuple[Any, int]:
    """
    Parses a single value (scalar, dict, or list) from tokens[pos].
    Returns (parsed_value, next_pos).

    For unquoted values that span multiple tokens before the next
    structural delimiter (e.g. 'infra bond'), joins them as a string.
    """
    if pos >= len(tokens):
        raise ValueError("Unexpected end of tokens while parsing value")

    token = tokens[pos]

    if token == '{':
        return _parse_dict_tokens(tokens, pos)

    if token == '[':
        return _parse_list_tokens(tokens, pos)

    if token.startswith('"'):
        # Quoted string — strip quotes and cast
        return _cast_scalar(_unquote(token)), pos + 1

    # Unquoted scalar — may span multiple words until next structural char
    parts = []
    while pos < len(tokens) and tokens[pos] not in ('{', '}', '[', ']', ',', ':'):
        parts.append(tokens[pos])
        pos += 1

    return _cast_scalar(" ".join(parts)), pos


# ── Scalar caster ──────────────────────────────────────────────────────────────

def _cast_scalar(value: str) -> Any:
    """
    Casts a scalar string value:
      - 'true' / 'false' (any case) → Python bool
      - Numbers (int or float)      → str  (never numeric)
      - Anything else               → str  (stripped)
    """
    stripped = value.strip()

    if stripped.lower() == "true":
        return True
    if stripped.lower() == "false":
        return False

    # Numbers → str (do not cast to int/float)
    try:
        float(stripped)
        return stripped   # it's a number — return as string
    except ValueError:
        pass

    return stripped


def _unquote(token: str) -> str:
    """Removes surrounding double quotes from a token if present."""
    if token.startswith('"') and token.endswith('"'):
        return token[1:-1]
    return token