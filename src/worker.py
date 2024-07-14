import os
import json
import asyncio
import logging

from io import BytesIO

from aio_pika import Message, connect

from utils import get_connection, doc_to_docx, docx_to_html
from html_to_json import html_to_json

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

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
                    try:
                        html, toc = docx_to_html(BytesIO(message.body))
                        converted = html_to_json(html)
                    except:
                        doc = BytesIO()
                        doc_to_docx(BytesIO(message.body), doc)
                        doc.seek(0)
                        html, toc = docx_to_html(doc)
                        converted = filter(lambda el: 
                                                not (el['content-type'] == 'text' 
                                                and 'evaluation copy of Aspose.Words' in el['content']), 
                                            json.loads(html_to_json(html)))
                        converted = json.dumps(list(converted))

                    await exchange.publish(
                            Message(
                                body=converted.encode(),
                                correlation_id=message.correlation_id
                            ),
                            routing_key=message.reply_to
                    )
                    logger.info("Task complete")

            except Exception as e:
                logging.exception("Processing error: "+str(e))

if __name__ == '__main__':
    asyncio.run(main())

