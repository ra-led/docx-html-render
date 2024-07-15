import os
import asyncio
import logging

from aio_pika import Message, connect
from typing import Annotated
from fastapi import FastAPI, File
from io import BytesIO, StringIO

from utils import get_connection, ConverterProxy

app = FastAPI()
logger = logging.getLogger(__name__)

converter = None

@app.post("/")
async def root(file: Annotated[bytes, File()]):
    global converter
    if converter is None:
        converter = ConverterProxy()

    return await converter.convert(file)