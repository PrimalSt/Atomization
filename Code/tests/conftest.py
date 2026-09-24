import httpx
import pytest

from src import direct_client


@pytest.fixture(autouse=True)
def no_real_network(monkeypatch):
    """Тесты не ходят в сеть: настоящий транспорт httpx падает, MockTransport работает."""

    def blocked(self, request: httpx.Request) -> httpx.Response:
        raise RuntimeError(f"Сетевой запрос в тестах запрещён: {request.method} {request.url}")

    monkeypatch.setattr(httpx.HTTPTransport, "handle_request", blocked)


@pytest.fixture(autouse=True)
def isolated_direct_token(monkeypatch, tmp_path):
    """Настоящий токен из окружения или .env проекта в тесты не попадает."""
    monkeypatch.delenv(direct_client.TOKEN_ENV_VAR, raising=False)
    monkeypatch.setattr(direct_client, "ENV_FILE", tmp_path / "absent.env")
