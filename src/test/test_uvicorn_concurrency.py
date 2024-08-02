import asyncio
import os

import aiohttp
import pytest

from test_utils import fake_worker, server, create_form_data

pytest_plugins = ('pytest_asyncio',)

@pytest.mark.asyncio
@pytest.mark.parametrize('server', [{}], indirect=['server'])
async def test_single_user(server, fake_worker):
    async with aiohttp.ClientSession(connector=aiohttp.TCPConnector(force_close=True)) as session:
        post = session.post('http://127.0.0.1:5000/', data=create_form_data())
        fake_worker.complete_tasks(count=1)
        async with post as resp:
            assert(resp.status == 200)
            assert(await resp.json() == 'completed by test')
            resp.close()

test_concurrency = 4
@pytest.mark.parametrize('requests_count, expected_success, expected_errors, server',[

    # no futures limit, concurrency limit = 4
    (test_concurrency-1,    test_concurrency-1, 0,                      {'concurrency_limit': test_concurrency, 'max_futures': '0'}),
    (test_concurrency,      0,                  test_concurrency,       {'concurrency_limit': test_concurrency, 'max_futures': '0'}),
    (test_concurrency+1,    0,                  test_concurrency+1,     {'concurrency_limit': test_concurrency, 'max_futures': '0'}),

    # max futures = 2, concurrency limit = 4, test around limits
    (1,                     1,                  0,                      {'concurrency_limit': test_concurrency, 'max_futures': '2'}),
    (2,                     2,                  0,                      {'concurrency_limit': test_concurrency, 'max_futures': '2'}),
    (3,                     2,                  1,                      {'concurrency_limit': test_concurrency, 'max_futures': '2'}),
    (test_concurrency-1,    2,                  test_concurrency-1-2,   {'concurrency_limit': test_concurrency, 'max_futures': '2'}),
    (test_concurrency,      0,                  test_concurrency,       {'concurrency_limit': test_concurrency, 'max_futures': '2'}),
    (test_concurrency+1,    0,                  test_concurrency+1,     {'concurrency_limit': test_concurrency, 'max_futures': '2'}),

    # max futures = 2, no concurrency limit, test around max futures and old concurrency limit
    (1,                     1,                  0,                      {'concurrency_limit': None, 'max_futures': '2'}),
    (2,                     2,                  0,                      {'concurrency_limit': None, 'max_futures': '2'}),
    (3,                     2,                  1,                      {'concurrency_limit': None, 'max_futures': '2'}),
    (test_concurrency-1,    2,                  test_concurrency-1-2,   {'concurrency_limit': None, 'max_futures': '2'}),
    (test_concurrency,      2,                  test_concurrency-2,     {'concurrency_limit': None, 'max_futures': '2'}),
    (test_concurrency+1,    2,                  test_concurrency+1-2,   {'concurrency_limit': None, 'max_futures': '2'}),

    # no limits, test around old limits
    (1,                     1,                  0,                      {'concurrency_limit': None, 'max_futures': '0'}),
    (2,                     2,                  0,                      {'concurrency_limit': None, 'max_futures': '0'}),
    (3,                     3,                  0,                      {'concurrency_limit': None, 'max_futures': '0'}),
    (test_concurrency-1,    test_concurrency-1, 0,                      {'concurrency_limit': None, 'max_futures': '0'}),
    (test_concurrency,      test_concurrency,   0,                      {'concurrency_limit': None, 'max_futures': '0'}),
    (test_concurrency+1,    test_concurrency+1, 0,                      {'concurrency_limit': None, 'max_futures': '0'}),

], indirect=['server'])
@pytest.mark.asyncio
@pytest.mark.flaky
async def test_concurrency_limit_boundaries(server, fake_worker, requests_count, expected_success, expected_errors):
    async with aiohttp.ClientSession(connector=aiohttp.TCPConnector(force_close=True)) as session:
        async def do_request():
            req = session.post('http://127.0.0.1:5000/', data=create_form_data())
            try:
                async with req as resp:
                    assert(resp.status == 200)
                    assert(await resp.json() == 'completed by test')
                    resp.close()
                    return True
            except (AssertionError, aiohttp.client_exceptions.ClientOSError):
                if req:
                    req.close()
                return False

        requests = [asyncio.create_task(do_request()) for _ in range(requests_count)]
        await asyncio.sleep(0.1)
        fake_worker.complete_tasks(count=requests_count)
        await asyncio.sleep(0.1)
        
        succeded = 0
        errored = 0
        results = await asyncio.gather(*requests)
        for res in results:
            if res: succeded += 1
            else: errored += 1

        assert(succeded == expected_success) # flaky
        assert(errored == expected_errors)
