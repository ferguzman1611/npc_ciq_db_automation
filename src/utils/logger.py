"""
Project-wide logger configuration.

All modules obtain their logger through get_logger(). Every run writes to a
single log file under logs/, named with the timestamp of when the process
started, and mirrors INFO-level output to the console.

The timestamp is captured once, at import time, so that every logger created
during the same run shares the same file. The file and console handlers are
also created once and reused across all loggers, which keeps a single open
file handle per run.
"""

import logging
from datetime import datetime

from config import LOGS_DIR

# Captured once when this module is first imported. Because imports are cached,
# every get_logger() call during the run reuses this same value.
_RUN_TIMESTAMP = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# Shared handler instances, built lazily on the first get_logger() call.
_handlers: list[logging.Handler] | None = None


def get_logger(name: str) -> logging.Logger:
    """
    Returns a logger that writes to the run's log file and the console.

    Args:
        name: Logger name, typically __name__ from the calling module. It is
              printed on every line so log entries can be traced back to the
              module that produced them.

    Returns:
        A configured logging.Logger instance.
    """
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # Attach the shared handlers only once per logger. propagate is disabled so
    # records are not also forwarded to the root logger, which would duplicate
    # every line.
    if not logger.handlers:
        for handler in _get_handlers():
            logger.addHandler(handler)
        logger.propagate = False

    return logger


def _get_handlers() -> list[logging.Handler]:
    """Builds the file and console handlers on first use and caches them."""
    global _handlers
    if _handlers is not None:
        return _handlers

    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOGS_DIR / f"{_RUN_TIMESTAMP}.log"

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # The file captures everything (DEBUG and up); the console shows INFO and up
    # to keep normal runs readable.
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    _handlers = [file_handler, console_handler]
    return _handlers
