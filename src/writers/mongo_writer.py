"""
Writes pipeline data to a MongoDB collection.

A dict upserts a single document (Global Attributes); a list either upserts
each document by a key or fully refreshes the collection (Regional / Network
Elements). Empty data leaves the collection untouched.

Writes can be disabled entirely with MONGODB_ENABLED=false in .env, in which
case this module logs a notice and returns immediately.
"""

from pymongo import MongoClient
from pymongo.errors import ConnectionFailure, OperationFailure

import config
from src.utils.logger import get_logger

logger = get_logger(__name__)

def write_to_mongo(
    data: dict | list[dict],
    collection_name: str,
    upsert_key: str | None = None,
) -> None:
    """
    Writes data to a MongoDB collection.

    Skips execution entirely if:
      - MONGODB_ENABLED is false in .env
      - data is an empty dict or empty list

    If data is a dict:
        Upserts a single document (replaces existing or inserts new).

    If data is a list of dicts:
        If upsert_key provided: upserts each document by that key.
        If upsert_key is None: drops collection and re-inserts all documents.

    Args:
        data:            A dict or list of dicts to write.
        collection_name: Target MongoDB collection name.
        upsert_key:      Field to use as upsert key for list writes.

    Raises:
        ConnectionFailure: If MongoDB is unreachable.
        OperationFailure:  If any DB operation fails.
        TypeError:         If data is neither a dict nor a list.
    """
    if not isinstance(data, (dict, list)):
        raise TypeError(f"data must be dict or list, got {type(data).__name__}")

    if not config.MONGODB_ENABLED:
        logger.info(f"MongoDB disabled; skipping write to '{collection_name}'.")
        return

    if not data:
        logger.warning(
            f"Data for '{collection_name}' is empty; skipping MongoDB write."
        )
        return

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


def _get_collection(client: MongoClient, collection_name: str):
    """Returns a MongoDB collection object from the configured database."""
    return client[config.MONGO_DB_NAME][collection_name]


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


def _write_many(
    collection,
    data: list[dict],
    collection_name: str,
    upsert_key: str | None = None,
) -> None:
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
                logger.warning(
                    f"Document missing upsert_key '{upsert_key}'; skipped: {doc}"
                )
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
            f"'{collection_name}': {upserted} inserted, {replaced} replaced "
            f"(upsert by '{upsert_key}')."
        )
    else:
        collection.drop()
        logger.debug(f"Collection '{collection_name}' dropped for full refresh.")
        result = collection.insert_many(data)
        logger.info(
            f"{len(result.inserted_ids)} documents inserted into '{collection_name}'."
        )