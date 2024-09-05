from enum import Enum
from loguru import logger


class Action(Enum):
    PASS = 'pass'
    UPDATE = 'update'
    REMOVE = 'remove'


def custom_callback(element):
    # Implement it
    logger.info('Custom callback started')
    ...
    action = Action.PASS
    return action, element