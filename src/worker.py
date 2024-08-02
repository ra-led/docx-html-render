import asyncio
from io import BytesIO
import os
from aio_pika import Message, connect
from loguru import logger
from doc_parse import doc_to_docx, docx_to_json
from utils import get_connection


async def process_message(message, exchange):
    logger.info(f"Received task (reply to: {message.reply_to})")
    
    try:
        logger.info("Starting conversion to JSON")
        converted = docx_to_json(BytesIO(message.body))
    except Exception as e:
        logger.exception("Error during direct conversion to JSON")
        logger.info("Starting conversion from DOC to DOCX")
        doc = BytesIO()
        try:
            doc_to_docx(BytesIO(message.body), doc)
        except Exception as e:
            logger.exception("Error during conversion from DOC to DOCX")
            return
        doc.seek(0)
        logger.info("Starting conversion from DOCX to JSON")
        try:
            converted = docx_to_json(doc)
        except Exception as e:
            logger.exception("Error during conversion from DOCX to JSON")
            return
    
    logger.info("Conversion completed")
    
    await exchange.publish(
        Message(
            body=converted.encode(),
            correlation_id=message.correlation_id
        ),
        routing_key=message.reply_to
    )
    logger.info("Message published back to exchange")
    logger.info("Task complete")


async def main():
    try:
        async with await get_connection() as connection:
            logger.info("Connection to RabbitMQ established")
            
            async with connection.channel() as channel:
                logger.info("Channel opened")
                
                exchange = channel.default_exchange
                
                queue_name = os.environ.get('CONVERTER_QUEUE', default='convert')
                async with channel.declare_queue(queue_name) as queue:
                    logger.info("Queue declared")
                    
                    logger.info("Waiting for tasks")
                    async with queue.iterator() as iterator:
                        async for message in iterator:
                            try:
                                async with message.process(requeue=False):
                                    await process_message(message, exchange)
                            except Exception as e:
                                logger.exception("Processing error")

    except Exception as e:
        logger.exception("Main error")

if __name__ == '__main__':
    asyncio.run(main())