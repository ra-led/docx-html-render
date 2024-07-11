import asyncio
import os
import uuid
import logging
import json

import aspose.words as aw
from aio_pika import Message, connect
from lxml import etree

logger = logging.getLogger(__name__)

async def get_connection():
    user = os.environ.get('RABBITMQ_USER', default='guest')
    pasw = os.environ.get('RABBITMQ_PASS', default='guest')
    host = os.environ.get('RABBITMQ_HOST', default='rabbitmq')
    port = os.environ.get('RABBITMQ_PORT', default=5672)

    return await connect(f'amqp://{user}:{pasw}@{host}:{port}')

def doc_to_docx(in_stream, out_stream):
    doc = aw.Document(in_stream)
    doc.save(out_stream, aw.SaveFormat.DOCX)

class ConverterProxy:

    def __init__(self):
        self.initialized = False
        self.futures = {}

    async def convert(self, data: bytes):
        if not self.initialized:
            self.initialized = True
            self.connection = await get_connection()
            self.channel = await self.connection.channel()
            self.callback_queue = await self.channel.declare_queue(exclusive=True)
            await self.callback_queue.consume(self.on_message, no_ack=True)

        correlation_id = str(uuid.uuid4())
        loop = asyncio.get_running_loop()
        future = loop.create_future()

        self.futures[correlation_id] = future

        await self.channel.default_exchange.publish(
            Message(
                data,
                correlation_id=correlation_id,
                reply_to=self.callback_queue.name
                ),
            routing_key=os.environ.get('CONVERTER_QUEUE', default='convert')
            )
        return await future

    async def on_message(self, message):
        if message.correlation_id is None:
            print(f"Bad message {message!r}")
            return

        future: asyncio.Future = self.futures.pop(message.correlation_id)
        future.set_result(message.body)

