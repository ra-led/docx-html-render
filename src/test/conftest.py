import pytest

def pytest_addoption(parser):
    parser.addoption("--host", action="store", default="127.0.0.1", help="Host IP address")
    parser.addoption("--port", action="store", default="8000", help="Port number")
    parser.addoption("--warning_threshold", action="store", default=0.9, type=float, help="Warning threshold")
    parser.addoption("--error_threshold", action="store", default=0.1, type=float, help="Error threshold")

@pytest.fixture
def host(request):
    return request.config.getoption("--host")

@pytest.fixture
def port(request):
    return request.config.getoption("--port")

@pytest.fixture
def warning_threshold(request):
    return request.config.getoption("--warning_threshold")

@pytest.fixture
def error_threshold(request):
    return request.config.getoption("--error_threshold")