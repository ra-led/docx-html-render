import asyncio
from io import BytesIO
import os
from aio_pika import Message, connect
from loguru import logger
from doc_parse import doc_to_docx, docx_to_json
from utils import get_connection


async def process_message(message, exchange):
    logger.info(f"Received task (reply to: {message.reply_to}, correlation_id: {message.correlation_id})")
    
    try:
        logger.info(f"Starting conversion to JSON (correlation_id: {message.correlation_id})")
        converted = docx_to_json(BytesIO(message.body))
    except Exception as e:
        logger.exception(f"Error during direct conversion to JSON (correlation_id: {message.correlation_id})")
        logger.info(f"Starting conversion from DOC to DOCX (correlation_id: {message.correlation_id})")
        doc = BytesIO()
        try:
            doc_to_docx(BytesIO(message.body), doc)
        except Exception as e:
            logger.exception(f"Error during conversion from DOC to DOCX (correlation_id: {message.correlation_id})")
            return
        doc.seek(0)
        logger.info(f"Starting conversion from DOCX to JSON (correlation_id: {message.correlation_id})")
        try:
            converted = docx_to_json(doc)
        except Exception as e:
            logger.exception(f"Error during conversion from DOCX to JSON (correlation_id: {message.correlation_id})")
            return
    
    logger.info(f"Conversion completed (correlation_id: {message.correlation_id})")
    
    await exchange.publish(
        Message(
            body=converted.encode(),
            correlation_id=message.correlation_id
        ),
        routing_key=message.reply_to
    )
    logger.info(f"Message published back to exchange (correlation_id: {message.correlation_id})")
    logger.info(f"Task complete (correlation_id: {message.correlation_id})")


async def main():
    try:
        async with await get_connection() as connection:
            logger.info("Connection to RabbitMQ established")
            
            async with connection.channel() as channel:
                logger.info("Channel opened")
                
                exchange = channel.default_exchange
                
                queue_name = os.environ.get('CONVERTER_QUEUE', default='convert')
                queue = await channel.declare_queue(queue_name)  # Await the coroutine
                logger.info("Queue declared")
                
                logger.info("Waiting for tasks")
                async with queue.iterator() as iterator:
                    async for message in iterator:
                        try:
                            async with message.process(requeue=False):
                                await process_message(message, exchange)
                        except Exception as e:
                            logger.exception(f"Processing error (correlation_id: {message.correlation_id})")

    except Exception as e:
        logger.exception("Main error")

if __name__ == '__main__':
    asyncio.run(main())