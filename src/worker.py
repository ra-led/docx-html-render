import asyncio
from io import BytesIO
import os
from aio_pika import Message, connect
from loguru import logger
from doc_parse import doc_to_docx, docx_to_json
from utils import get_connection


async def main():
    try:
        connection = await get_connection()
        logger.info("Connection to RabbitMQ established")
        
        channel = await connection.channel()
        logger.info("Channel opened")
        
        exchange = channel.default_exchange
        
        queue = await channel.declare_queue(os.environ.get('CONVERTER_QUEUE', default='convert'))
        logger.info("Queue declared")
        
        logger.info("Waiting for tasks")
        async with queue.iterator() as iterator:
            async for message in iterator:
                try:
                    async with message.process(requeue=False):
                        logger.info(f"Received task (reply to: {message.reply_to})")
                        
                        try:
                            logger.info("Starting conversion to JSON")
                            converted = docx_to_json(BytesIO(message.body))
                        except:
                            logger.info("Starting conversion from DOC to DOCX")
                            doc = BytesIO()
                            doc_to_docx(BytesIO(message.body), doc)
                            doc.seek(0)
                            logger.info("Starting conversion from DOCX to JSON")
                            converted = docx_to_json(doc)
                        
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

                except Exception as e:
                    logger.exception("Processing error: " + str(e))

    except Exception as e:
        logger.exception("Main error: " + str(e))

if __name__ == '__main__':
    asyncio.run(main())
