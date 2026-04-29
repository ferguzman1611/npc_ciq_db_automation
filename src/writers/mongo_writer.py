"""
src/writers/mongo_writer.py
----------------------------
Writes data to a MongoDB collection.

Supports two modes:
  - dict   → inserts/replaces a single document (used by Global Attributes)
  - list   → drops the collection and inserts all documents fresh (used by Regional Attributes)

The list strategy (drop + insert_many) is intentional: regional data is
always written as a complete dataset, so replacing individual documents
by region would require an extra filter key. A full refresh is simpler,
faster, and keeps the collection consistent with the Excel at all times.
"""

from pymongo import MongoClient
from pymongo.errors import ConnectionFailure, OperationFailure

import config
from src.utils.logger import get_logger

logger = get_logger(__name__)


def _get_collection(client: MongoClient, collection_name: str):
    """Returns a MongoDB collection object from the configured database."""
    return client[config.MONGO_DB_NAME][collection_name]


def write_to_mongo(
    data: dict | list[dict],
    collection_name: str,
    upsert_key: str | None = None,
) -> None:
    """
    Writes data to a MongoDB collection.

    If data is a dict:
        Replaces the single existing document in the collection (upsert).
        Idempotent — running twice does not create duplicates.

    If data is a list of dicts:
        If upsert_key is provided: upserts each document individually by that key.
        If upsert_key is None: drops the collection and re-inserts all documents.

    Args:
        data:            A dict (single document) or list of dicts (multiple documents).
        collection_name: Name of the target MongoDB collection.

    Raises:
        ConnectionFailure: If MongoDB is unreachable.
        OperationFailure:  If any DB operation fails.
        TypeError:         If data is neither a dict nor a list.
    """
    if not isinstance(data, (dict, list)):
        raise TypeError(f"data must be dict or list, got {type(data).__name__}")

    logger.info(f"Connecting to MongoDB at {config.MONGO_URI} ...")

    try:
        client = MongoClient(config.MONGO_URI, serverSelectionTimeoutMS=5000)
        client.admin.command("ping")
        logger.info("MongoDB connection successful.")
    except ConnectionFailure as exc:
        logger.error(f"Could not connect to MongoDB: {exc}")
        raise

    try:
        collection = _get_collection(client, collection_name)

        if isinstance(data, dict):
            _write_single(collection, data, collection_name)
        else:
            _write_many(collection, data, collection_name, upsert_key)

    except OperationFailure as exc:
        logger.error(f"MongoDB operation failed: {exc}")
        raise
    finally:
        client.close()


def _write_single(collection, data: dict, collection_name: str) -> None:
    """Upserts a single document, replacing whatever is currently in the collection."""
    result = collection.replace_one(
        filter={},
        replacement=data,
        upsert=True,
    )
    if result.upserted_id:
        logger.info(f"New document inserted in '{collection_name}'.")
    else:
        logger.info(f"Existing document replaced in '{collection_name}'.")


def _write_many(collection, data: list[dict], collection_name: str, upsert_key: str | None = None) -> None:
    """
    If upsert_key is given: upserts each document by that key (e.g. 'Node_Name').
    Otherwise: drops the collection and re-inserts all documents fresh.
    """
    if upsert_key:
        upserted = 0
        replaced = 0
        for doc in data:
            key_value = doc.get(upsert_key)
            if key_value is None:
                logger.warning(f"Document missing upsert_key '{upsert_key}' — skipped: {doc}")
                continue
            result = collection.replace_one(
                filter={upsert_key: key_value},
                replacement=doc,
                upsert=True,
            )
            if result.upserted_id:
                upserted += 1
            else:
                replaced += 1
        logger.info(
            f"'{collection_name}': {upserted} inserted, {replaced} replaced (upsert by '{upsert_key}')."
        )
    else:
        collection.drop()
        logger.debug(f"Collection '{collection_name}' dropped for full refresh.")
        result = collection.insert_many(data)
        logger.info(
            f"{len(result.inserted_ids)} documents inserted into '{collection_name}'."
        )