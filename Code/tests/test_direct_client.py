import datetime as dt
import json
import logging

import httpx
import pytest

from src import direct_client
from src.config_loader import load_report_config
from src.direct_client import (
    BASE_FIELDS,
    MAX_ATTEMPTS,
    REPORTS_URL,
    TOKEN_ENV_VAR,
    DirectApiError,
    DirectAuthError,
    DirectLimitError,
    DirectReportTimeout,
    build_report_request,
    fetch_campaign_report,
    resolve_token,
)
from src.schemas import DirectApiConfig, ReportConfig

TODAY = dt.date(2026, 9, 24)
TOKEN = "test-token-123"
TSV = (
    "Date\tCampaignId\tCampaignName\tImpressions\tClicks\tCost\tConversions\n"
    "2026-09-22\t98743239\tПоиск_РФ_Бренд\t4126\t158\t5964.74\t8\n"
    "2026-09-23\t98743239\tПоиск_РФ_Бренд\t3955\t132\t5222.40\t--\n"
    "2026-09-23\t72731077\tРСЯ_РФ_Интересы\t17734\t350\t4303.17\t10\n"
)


class FakeDirect:
    """Отвечает заготовленными ответами по очереди и запоминает запросы."""

    def __init__(self, *responses: httpx.Response) -> None:
        self.responses = list(responses)
        self.requests: list[httpx.Request] = []

    def handle(self, request: httpx.Request) -> httpx.Response:
        self.requests.append(request)
        return self.responses.pop(0)

    def client(self) -> httpx.Client:
        return httpx.Client(transport=httpx.MockTransport(self.handle))

    def body(self, index: int = 0) -> dict:
        return json.loads(self.requests[index].content)["params"]


@pytest.fixture
def config() -> ReportConfig:
    return load_report_config()


def with_direct(config: ReportConfig, **settings) -> ReportConfig:
    return config.model_copy(update={"direct_api": DirectApiConfig(**settings)})


def fetch(config: ReportConfig, fake: FakeDirect, sleeps: list[float] | None = None, **kwargs):
    recorder = sleeps if sleeps is not None else []
    return fetch_campaign_report(
        config, token=TOKEN, today=TODAY, client=fake.client(), sleep=recorder.append, **kwargs
    )


def api_error(status: int, code: int, text: str) -> httpx.Response:
    return httpx.Response(
        status,
        json={"error": {"request_id": "req-1", "error_code": str(code), "error_string": text, "error_detail": "Подробности"}},
    )


# --- Успешный запрос ------------------------------------------------------------


def test_success_on_first_try_sends_spec_request_and_parses_records(config):
    fake = FakeDirect(httpx.Response(200, text=TSV, headers={"RequestId": "req-ok"}))
    sleeps: list[float] = []

    records = fetch(config, fake, sleeps)

    [request] = fake.requests
    assert request.method == "POST" and str(request.url) == REPORTS_URL
    assert request.headers["Authorization"] == f"Bearer {TOKEN}"
    assert request.headers["returnMoneyInMicros"] == "false"
    assert request.headers["skipReportHeader"] == "true"
    assert request.headers["skipReportSummary"] == "true"
    assert request.headers["processingMode"] == "auto"
    assert "Client-Login" not in request.headers
    params = fake.body()
    assert params["ReportType"] == "CAMPAIGN_PERFORMANCE_REPORT"
    assert params["DateRangeType"] == "CUSTOM_DATE"
    assert params["Format"] == "TSV"
    assert params["IncludeVAT"] == "YES"
    assert params["FieldNames"] == [*BASE_FIELDS, "Conversions"]
    assert "Goals" not in params
    assert sleeps == []

    assert len(records) == 3
    first = records[0]
    assert (first.date, first.campaign_id, first.campaign_name) == (dt.date(2026, 9, 22), 98743239, "Поиск_РФ_Бренд")
    assert (first.impressions, first.clicks, first.cost, first.conversions) == (4126, 158, 5964.74, 8)
    assert records[1].conversions == 0  # «--» — нет данных


def test_client_login_and_vat_settings(config):
    fake = FakeDirect(httpx.Response(200, text=TSV))

    fetch(with_direct(config, client_login="delivery-plus", include_vat=False), fake)

    assert fake.requests[0].headers["Client-Login"] == "delivery-plus"
    assert fake.body()["IncludeVAT"] == "NO"


def test_goals_request_and_sum_per_goal_columns(config):
    tsv = (
        "Date\tCampaignId\tCampaignName\tImpressions\tClicks\tCost\tConversions_111_LSC\tConversions_222_LSC\n"
        "2026-09-23\t98743239\tПоиск_РФ_Бренд\t3955\t132\t5222.40\t3\t--\n"
        "2026-09-23\t72731077\tРСЯ_РФ_Интересы\t17734\t350\t4303.17\t4\t2\n"
    )
    fake = FakeDirect(httpx.Response(200, text=tsv))

    records = fetch(with_direct(config, goals=[111, 222]), fake)

    params = fake.body()
    assert params["Goals"] == ["111", "222"]
    assert "Conversions" in params["FieldNames"]  # с Goals Директ делит это поле по целям
    assert [record.conversions for record in records] == [3, 6]


def test_missing_goal_columns_is_an_error(config):
    fake = FakeDirect(httpx.Response(200, text=TSV))

    with pytest.raises(DirectApiError, match="нет колонок конверсий для целей: 111"):
        fetch(with_direct(config, goals=[111]), fake)


def test_empty_report_gives_no_records(config):
    fake = FakeDirect(httpx.Response(200, text="Date\tCampaignId\tCampaignName\tImpressions\tClicks\tCost\tConversions\n"))

    assert fetch(config, fake) == []


# --- Очередь отчётов ------------------------------------------------------------


def test_queue_201_202_then_200(config, caplog):
    fake = FakeDirect(
        httpx.Response(201, headers={"retryIn": "1", "reportsInQueue": "2"}),
        httpx.Response(202, headers={"retryIn": "1"}),
        httpx.Response(200, text=TSV),
    )
    sleeps: list[float] = []

    with caplog.at_level(logging.INFO, logger="src"):
        records = fetch(config, fake, sleeps)

    assert len(records) == 3
    assert sleeps == [1.0, 1.0]
    assert len(fake.requests) == 3
    assert fake.body(0) == fake.body(1) == fake.body(2)  # тот же отчёт, а не новый в очереди
    assert "поставлен в очередь" in caplog.text and "формируется" in caplog.text
    assert "отчётов в очереди: 2" in caplog.text
    assert "скачано строк: 3" in caplog.text


def test_default_retry_in_is_five_seconds(config):
    fake = FakeDirect(httpx.Response(201), httpx.Response(200, text=TSV))
    sleeps: list[float] = []

    fetch(config, fake, sleeps)

    assert sleeps == [5.0]


def test_report_not_ready_after_max_attempts(config):
    fake = FakeDirect(*[httpx.Response(202, headers={"retryIn": "1"}) for _ in range(MAX_ATTEMPTS)])
    sleeps: list[float] = []

    with pytest.raises(DirectReportTimeout, match=f"не готов после {MAX_ATTEMPTS} попыток"):
        fetch(config, fake, sleeps)

    assert len(fake.requests) == MAX_ATTEMPTS
    assert len(sleeps) == MAX_ATTEMPTS - 1  # после последней попытки не ждём


def test_server_error_is_retried(config, caplog):
    fake = FakeDirect(api_error(500, 1000, "Внутренняя ошибка сервера"), httpx.Response(200, text=TSV))

    with caplog.at_level(logging.WARNING, logger="src"):
        records = fetch(config, fake)

    assert len(records) == 3
    assert "временная ошибка сервера" in caplog.text


# --- Ошибки ---------------------------------------------------------------------


@pytest.mark.parametrize(
    "response",
    [
        httpx.Response(401, text="Unauthorized"),
        httpx.Response(403, text="Forbidden"),
        api_error(400, 53, "Ошибка авторизации"),
        api_error(400, 513, "Ваш логин не подключен к Яндекс.Директу"),
    ],
)
def test_auth_errors(config, response, caplog):
    fake = FakeDirect(response)

    with caplog.at_level(logging.DEBUG, logger="src"), pytest.raises(DirectAuthError) as error:
        fetch(config, fake)

    assert "Ошибка авторизации в API Директа" in str(error.value)
    assert TOKEN_ENV_VAR in str(error.value)
    assert TOKEN not in str(error.value) + caplog.text  # токен не утекает в сообщения и лог
    assert len(fake.requests) == 1  # авторизацию не повторяем


@pytest.mark.parametrize(
    ("response", "code"),
    [
        (api_error(400, 152, "Недостаточно баллов"), 152),
        (api_error(400, 9000, "Превышен лимит на количество отчетов в очереди"), 9000),
        (api_error(400, 56, "Превышен лимит запросов"), 56),
        (httpx.Response(429, text="Too Many Requests"), None),
    ],
)
def test_limit_errors(config, response, code):
    fake = FakeDirect(response)

    with pytest.raises(DirectLimitError) as error:
        fetch(config, fake)

    assert error.value.code == code
    assert str(error.value).startswith("Лимит API Директа")


def test_request_error_carries_details(config):
    fake = FakeDirect(api_error(400, 8000, "Неверный запрос"))

    with pytest.raises(DirectApiError) as error:
        fetch(config, fake)

    assert type(error.value) is DirectApiError
    assert str(error.value) == "Ошибка API Директа: Неверный запрос. Подробности (HTTP 400, код 8000, RequestId req-1)"


def test_network_failure_is_retried_then_reported(config):
    def always_down(request: httpx.Request) -> httpx.Response:
        raise httpx.ConnectError("connection refused", request=request)

    client = httpx.Client(transport=httpx.MockTransport(always_down))
    sleeps: list[float] = []

    with pytest.raises(DirectApiError, match="Нет связи с API Директа"):
        fetch_campaign_report(config, token=TOKEN, today=TODAY, client=client, sleep=sleeps.append, max_attempts=3)

    assert sleeps == [5.0, 5.0]


# --- Токен ----------------------------------------------------------------------


def test_missing_token_raises_auth_error(config):
    with pytest.raises(DirectAuthError, match=TOKEN_ENV_VAR):
        fetch_campaign_report(config, today=TODAY, client=FakeDirect().client())


def test_token_sources_priority(monkeypatch, tmp_path):
    env_file = tmp_path / ".env"
    env_file.write_text(f'# секреты\n{TOKEN_ENV_VAR}="from-file"\n', encoding="utf-8")
    monkeypatch.setattr(direct_client, "ENV_FILE", env_file)

    assert resolve_token() == "from-file"
    monkeypatch.setenv(TOKEN_ENV_VAR, "from-env")
    assert resolve_token() == "from-env"
    assert resolve_token("explicit") == "explicit"


# --- Период CUSTOM_DATE ---------------------------------------------------------


def test_custom_date_range_follows_project_period(config):
    request = build_report_request(config, today=TODAY)

    params = request.body["params"]
    assert params["DateRangeType"] == "CUSTOM_DATE"
    # period_days: 30, конец — вчера: сегодняшняя статистика неполная.
    assert params["SelectionCriteria"] == {"DateFrom": "2026-08-25", "DateTo": "2026-09-23"}
    assert (request.date_to - request.date_from).days + 1 == config.report_metadata.period_days


def test_custom_date_range_respects_explicit_period_end(config):
    metadata = config.report_metadata.model_copy(update={"period_days": 31, "date_to": dt.date(2026, 8, 31)})

    request = build_report_request(config.model_copy(update={"report_metadata": metadata}), today=TODAY)

    assert request.body["params"]["SelectionCriteria"] == {"DateFrom": "2026-08-01", "DateTo": "2026-08-31"}


def test_report_name_is_stable_and_depends_on_parameters(config):
    first = build_report_request(config, today=TODAY)
    again = build_report_request(config, today=TODAY)
    other = build_report_request(with_direct(config, include_vat=False), today=TODAY)

    assert first.name == again.name
    assert first.name != other.name
    assert len(first.name) <= 255
