"""
src/writers/json_writer.py
--------------------------
Writes a Python dictionary to a JSON file in the output/ directory.
"""

import json
from pathlib import Path

from src.utils.logger import get_logger

logger = get_logger(__name__)


def write_json(data: dict, output_path: Path) -> None:
    """
    Serializes a dictionary to a formatted JSON file.

    Creates the output directory if it does not exist.

    Args:
        data:        Dictionary to serialize.
        output_path: Full path (including filename) for the output file.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)

    logger.info(f"Writing JSON to: {output_path}")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

    logger.info(f"JSON file written successfully ({len(data)} keys).")
