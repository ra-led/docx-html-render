import asyncio
import os
import uuid
import aspose.words as aw
from aio_pika import Message, connect
import docx
from loguru import logger
from doc_parse import DocHandler, DocHTML


async def get_connection():
    """
    Establishes a connection to the RabbitMQ server.

    Returns:
        aio_pika.Connection: The connection object to the RabbitMQ server.
    """
    user = os.environ.get('RABBITMQ_USER', default='guest')
    pasw = os.environ.get('RABBITMQ_PASS', default='guest')
    host = os.environ.get('RABBITMQ_HOST', default='rabbitmq')
    port = os.environ.get('RABBITMQ_PORT', default=5672)

    return await connect(f'amqp://{user}:{pasw}@{host}:{port}')


class ConverterProxy:
    """
    A proxy class to handle document conversion requests via RabbitMQ.
    """

    def __init__(self):
        """
        Initializes the ConverterProxy instance.
        """
        self.initialized = False
        self.futures = {}

    async def convert(self, data: bytes):
        """
        Sends a document conversion request to the RabbitMQ queue and waits for the response.

        Args:
            data (bytes): The document data to be converted.

        Returns:
            bytes: The converted document data.
        """
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
        """
        Handles incoming messages from the RabbitMQ callback queue.

        Args:
            message (aio_pika.IncomingMessage): The incoming message from the RabbitMQ queue.
        """
        if message.correlation_id is None:
            logger.error(f"Bad message {message!r}")
            return

        future: asyncio.Future = self.futures.pop(message.correlation_id)
        future.set_result(message.body)


def doc_to_docx(in_stream, out_stream):
    """
    Converts a .doc file to a .docx file using Aspose.Words.

    Args:
        in_stream (io.BytesIO): The input stream containing the .doc file.
        out_stream (io.BytesIO): The output stream to write the .docx file.
    """
    doc = aw.Document(in_stream)
    doc.save(out_stream, aw.SaveFormat.DOCX)


def docx_to_html(docx_path: str) -> tuple:
    """
    Converts a DOCX document to HTML.
    
    Args:
        docx_path (str): The path to the DOCX file.
    
    Returns:
        tuple: A tuple containing the HTML content and table of contents links.
    """
    doc = docx.Document(docx_path)
    handler = DocHandler(doc)
    converter = DocHTML()
    
    return converter.get_html(handler)
