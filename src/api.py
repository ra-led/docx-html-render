import os
import asyncio
import logging

from aio_pika import Message, connect
from typing import Annotated
from fastapi import FastAPI, File, HTTPException
from io import BytesIO, StringIO

from utils import get_connection, ConverterProxy, FuturesLimitReachedException

def create_app():
    app = FastAPI()
    app.converter = ConverterProxy()
    logger = logging.getLogger(__name__)

    @app.post("/")
    async def root(file: Annotated[bytes, File()]):
        try:
            return await app.converter.convert(file)
        except FuturesLimitReachedException:
            raise HTTPException(status_code=429, detail='Too many requests in progress, try later')
    
    return app
