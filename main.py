"""
Entry point for the CIQ pipeline.

Usage:
    python main.py --global              # Global Attributes only
    python main.py --regional            # Regional Attributes only
    python main.py --global --regional   # both
    python main.py                       # all pipelines (default)
"""

import argparse
import sys

import config
from src.readers.global_reader import read_global_attributes
from src.readers.regional_reader import read_regional_attributes
from src.readers.network_elements_reader import read_network_elements
from src.writers.json_writer import write_json
from src.writers.mongo_writer import write_to_mongo
from src.utils.logger import get_logger

logger = get_logger(__name__)


def run_global_attributes() -> None:
    """Runs the full pipeline for the Global Attributes sheet."""
    logger.info("Starting Global Attributes pipeline")
    data = read_global_attributes()
    write_json(data, config.OUTPUT_FILE_GLOBAL)
    write_to_mongo(data, config.MONGO_COLLECTION_GLOBAL)
    logger.info("Global Attributes pipeline completed")


def run_regional_attributes() -> None:
    """Runs the full pipeline for the Regional Attributes sheet."""
    logger.info("Starting Regional Attributes pipeline")
    data = read_regional_attributes()
    write_json(data, config.OUTPUT_FILE_REGIONAL)
    write_to_mongo(data, config.MONGO_COLLECTION_REGIONAL)
    logger.info("Regional Attributes pipeline completed")


def run_network_elements() -> None:
    """Runs the full pipeline for all *-InputTable sheets."""
    logger.info("Starting Network Elements pipeline")
    data = read_network_elements()
    write_json(data, config.OUTPUT_FILE_NETWORK)
    write_to_mongo(data, config.MONGO_COLLECTION_NETWORK, upsert_key="Node_Name")
    logger.info("Network Elements pipeline completed")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Parse CIQ Excel sheets and write the results to JSON and MongoDB."
    )
    parser.add_argument(
        "--global",
        dest="run_global",
        action="store_true",
        help="Run the Global Attributes pipeline.",
    )
    parser.add_argument(
        "--regional",
        dest="run_regional",
        action="store_true",
        help="Run the Regional Attributes pipeline.",
    )
    parser.add_argument(
        "--network",
        dest="run_network",
        action="store_true",
        help="Run the Network Elements pipeline.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    config.LOGS_DIR.mkdir(parents=True, exist_ok=True)

    args = parse_args()

    # With no flags, run every pipeline.
    run_all = not args.run_global and not args.run_regional and not args.run_network

    try:
        if args.run_global or run_all:
            run_global_attributes()

        if args.run_regional or run_all:
            run_regional_attributes()

        if args.run_network or run_all:
            run_network_elements()

    except FileNotFoundError as exc:
        logger.error(f"Input file error: {exc}")
        sys.exit(1)
    except ValueError as exc:
        logger.error(f"Parsing error: {exc}")
        sys.exit(1)
    except Exception as exc:
        logger.error(f"Unexpected error: {exc}", exc_info=True)
        sys.exit(1)