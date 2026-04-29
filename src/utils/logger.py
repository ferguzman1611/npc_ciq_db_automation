"""
src/utils/logger.py
-------------------
Configures the project-wide logger.
Creates a timestamped log file in the logs/ directory
and also streams output to the console.
"""

import logging
from datetime import datetime
from config import LOGS_DIR


def get_logger(name: str) -> logging.Logger:
    """
    Returns a configured logger instance.

    Each run creates a new log file named YYYY-MM-DD_HH-MM-SS.log
    inside the logs/ directory. Console output mirrors the file.

    Args:
        name: Typically __name__ from the calling module.

    Returns:
        A configured logging.Logger instance.
    """
    LOGS_DIR.mkdir(parents=True, exist_ok=True)

    timestamp  = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_file   = LOGS_DIR / f"{timestamp}.log"

    logger     = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # Avoid adding duplicate handlers if get_logger is called multiple times
    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # File handler
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger
