import asyncio
import os
import uuid
from aio_pika import Message, connect
from loguru import logger


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


class FuturesLimitReachedException(Exception):
    pass

class ConverterProxy:
    """
    A proxy class to handle document conversion requests via RabbitMQ.
    """

    def __init__(self):
        """
        Initializes the ConverterProxy instance.
        """
        self.initialized = False
        self.initializing = False
        self.futures = {}
        self.futures_limit = int(os.environ.get('MAX_CONVERTER_FUTURES', default='0'))

    async def convert(self, data: bytes):
        """
        Sends a document conversion request to the RabbitMQ queue and waits for the response.

        Args:
            data (bytes): The document data to be converted.

        Returns:
            bytes: The converted document data.
        """
        if not self.initialized:
            self.initializing = True
            logger.info("Initializing ConverterProxy...")
            self.connection = await get_connection()
            self.channel = await self.connection.channel()
            self.callback_queue = await self.channel.declare_queue(exclusive=True)
            await self.callback_queue.consume(self.on_message, no_ack=True)
            self.initialized = True
            self.initializing = False
            logger.info("ConverterProxy initialized.")

        while self.initializing:
            await asyncio.sleep(0.1)

        if self.futures_limit > 0 and len(self.futures) >= self.futures_limit:
            logger.error("Futures limit reached.")
            raise FuturesLimitReachedException()

        correlation_id = str(uuid.uuid4())
        loop = asyncio.get_running_loop()
        future = loop.create_future()

        self.futures[correlation_id] = future

        logger.info(f"Sending conversion request with correlation_id: {correlation_id}")
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

        logger.info(f"Received message with correlation_id: {message.correlation_id}")
        future: asyncio.Future = self.futures.pop(message.correlation_id)
        future.set_result(message.body)
