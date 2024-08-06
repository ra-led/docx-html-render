import os
import sys
import contextlib
import threading
import time

import asyncio
import aiohttp
import pytest
import uvicorn
from aio_pika import Message
from aio_pika.exceptions import QueueEmpty
from loguru import logger

sys.path.append('..')
from utils import get_connection

queue_name = os.environ.get('CONVERTER_QUEUE', default='convert')

class Server(uvicorn.Server):
    def install_signal_handlers(self):
        pass

    @contextlib.contextmanager
    def run_in_thread(self, args = tuple(), kwargs = {}):
        thread = threading.Thread(target=self.run)
        thread.start()
        try:
            while not self.started:
                time.sleep(1e-3)
            yield
        finally:
            self.should_exit = True
            thread.join()

async def purge_converter_queue():
    connection = await get_connection()
    logger.debug('rabbit connected')
    channel = await connection.channel()
    queue = await channel.declare_queue(queue_name)
    await queue.purge()
    logger.debug('queue purged')

@pytest.fixture(scope="function")
def server(request):
    concurrency_limit = request.param.get('concurrency_limit', 4)
    os.environ['MAX_CONVERTER_FUTURES'] = str(request.param.get('max_futures', 0))
    connected = False
    start = time.time()
    while not connected:
        try:
            conn = asyncio.run(purge_converter_queue())
            connected = True
        except:
            time.sleep(1)
            assert time.time() - start < 30

    if concurrency_limit is None:
        server = Server(uvicorn.Config("api:create_app", port = 5000, host = '0.0.0.0'))
    else:
        server = Server(uvicorn.Config("api:create_app", limit_concurrency = concurrency_limit, port = 5000, host = '0.0.0.0'))
    with server.run_in_thread():
        logger.debug('server started '+str(time.time()))
        yield
        asyncio.run(asyncio.sleep(0)) # https://docs.aiohttp.org/en/stable/client_advanced.html#graceful-shutdown
        logger.debug('server finished '+str(time.time()))

class FakeWorker(threading.Thread):

    def __init__(self):
        self._stop_waiting = False
        self.count = 0
        threading.Thread.__init__(self)

    def run(self):
        logger.debug("Fake worker start")
        asyncio.run(self.process())

    def stop(self):
        self._stop_waiting = True
        self.join()
        logger.debug("Fake worker stop")

    def complete_tasks(self, count):
        self.count = count
        self.start()

    async def process(self):
        try:
            logger.debug('fake worker starting')
            connection = await get_connection()
            channel = await connection.channel()
            exchange = channel.default_exchange
            queue = await channel.declare_queue(queue_name)
            logger.debug('waiting tasks (count='+str(self.count)+')')
            for _ in range(self.count):
                msg = None
                while msg is None:
                    if self._stop_waiting: 
                        logger.debug("Fake worker stop condition reached")
                        return
                    try:
                        msg = await queue.get()
                    except QueueEmpty:
                        await asyncio.sleep(0.1)
                await exchange.publish(
                        Message(
                            body=b'completed by test',
                            correlation_id=msg.correlation_id
                            ),
                            routing_key=msg.reply_to
                        )
                await msg.ack()
                logger.debug('task complete')
            await connection.close()
        except Exception as e:
            logger.exception(e)

@pytest.fixture(scope="function")
def fake_worker():
    instance = FakeWorker()
    yield instance
    instance.stop()

def create_form_data():
    data = aiohttp.FormData()
    data.add_field('file', 
            b'test data bytes', # for FakeWorker any data will be fine
            filename='doc.doc',
            )
    return data
