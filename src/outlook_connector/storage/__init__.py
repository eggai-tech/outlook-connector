# TODO Implement
# from outlook_connector.schemas import Email

from structlog import get_logger

logger = get_logger()


class StorageError(Exception):
    pass


def save_to_storage(message):
    logger.debug("Saving to storage...", message_id=message.id)
