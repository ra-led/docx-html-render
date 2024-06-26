import os
import asyncio
import logging

from io import BytesIO

from aio_pika import Message, connect

from app import docx_to_html
from utils import get_connection

logger = logging.getLogger(__name__)

async def main():
    connection = await get_connection()
    channel = await connection.channel()
    exchange = channel.default_exchange

    queue = await channel.declare_queue(os.environ.get('CONVERTER_QUEUE', 
        default='convert'))

    logging.info("Waiting for tasks")
    async with queue.iterator() as iterator:
        async for message in iterator:
            try:
                async with message.process(requeue=False):
                    logger.info(f"Received task (reply to: {message.reply_to})")
                    html, toc = docx_to_html(BytesIO(message.body))
                    await exchange.publish(
                            Message(
                                body=html.encode(),
                                correlation_id=message.correlation_id
                            ),
                            routing_key=message.reply_to
                    )
                    logger.info("Task complete")

            except Exception as e:
                logging.exception("Processing error: "+str(e))

if __name__ == '__main__':
    asyncio.run(main())

