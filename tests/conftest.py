from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app.config import ensure_directories
from app.container import store
from app.main import app


@pytest.fixture(autouse=True)
def _reset_state():
    ensure_directories()
    store._sessions.clear()
    yield
    store._sessions.clear()


@pytest.fixture
def client():
    with TestClient(app) as test_client:
        yield test_client


@pytest.fixture
def documents_dir() -> Path:
    from app.config import settings

    return settings.documents_dir


@pytest.fixture
def templates_dir() -> Path:
    from app.config import settings

    return settings.templates_dir
