import json
import logging
import sys
from typing import Annotated
from fastapi import FastAPI, File, HTTPException
from loguru import logger

from utils import ConverterProxy, FuturesLimitReachedException


class InterceptHandler(logging.Handler):
    def emit(self, record):
        # Get corresponding Loguru level if it exists
        try:
            level = logger.level(record.levelname).name
        except ValueError:
            level = record.levelno

        # Find caller from where originated the logged message
        frame, depth = logging.currentframe(), 2
        while frame.f_code.co_filename == logging.__file__:
            frame = frame.f_back
            depth += 1

        logger.opt(depth=depth, exception=record.exc_info).log(level, record.getMessage())


def setup_logging():
    # Intercept everything at the root logger
    logging.root.handlers = [InterceptHandler()]
    logging.root.setLevel(logging.INFO)

    # Remove every other logger's handlers and propagate to root logger
    for name in logging.root.manager.loggerDict.keys():
        logging.getLogger(name).handlers = []
        logging.getLogger(name).propagate = True

    # Configure Loguru
    logger.configure(handlers=[{
        "sink": sys.stdout, "level": logging.getLevelName(logging.root.level)
    }])
    # Rotates the log file when it reaches 500 MB
    logger.add("app.log", rotation="500 MB")
    # Logs errors and above to a separate file
    logger.add("errors.log", level="ERROR")


def create_app():
    app = FastAPI()
    app.converter = ConverterProxy()

    # Setup logging for Uvicorn
    setup_logging()

    @app.post("/")
    async def root(file: Annotated[bytes, File()]):
        try:
            result = await app.converter.convert(file)
            return json.loads(result.decode())
        except FuturesLimitReachedException:
            logger.error("Too many requests in progress, try later")
            raise HTTPException(status_code=429, detail='Too many requests in progress, try later')
    
    return app
