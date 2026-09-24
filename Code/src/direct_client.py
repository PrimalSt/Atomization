"""Клиент Reports API Яндекс Директа (v5): суточная статистика кампаний → DailyAdRecord.

Запрос уходит в JSON-эндпоинт ``/json/v5/reports`` (адрес ``/v5/reports`` принимает
только XML). Режим processingMode=auto: готовый отчёт приходит сразу (HTTP 200),
иначе Директ ставит его в очередь (201) или формирует (202) — тогда тот же запрос
повторяется через ``retryIn`` секунд.
"""

import csv
import datetime as dt
import hashlib
import io
import json
import logging
import os
import re
import time
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import httpx

from src.config_loader import PROJECT_ROOT
from src.data_loader import parse_ad_rows
from src.schemas import DailyAdRecord, DirectApiConfig, ReportConfig

logger = logging.getLogger(__name__)

REPORTS_URL = "https://api.direct.yandex.com/json/v5/reports"
SANDBOX_REPORTS_URL = "https://api-sandbox.direct.yandex.com/json/v5/reports"
TOKEN_ENV_VAR = "YANDEX_DIRECT_TOKEN"
ENV_FILE = PROJECT_ROOT / ".env"
REPORT_TYPE = "CAMPAIGN_PERFORMANCE_REPORT"
BASE_FIELDS = ("Date", "CampaignId", "CampaignName", "Impressions", "Clicks", "Cost")
MAX_ATTEMPTS = 12
DEFAULT_RETRY_IN = 5.0
SOURCE_NAME = "API Директа"

_TIMEOUT = httpx.Timeout(30.0, read=180.0)  # онлайн-отчёт Директ может формировать пару минут
_AUTH_ERROR_CODES = frozenset({53, 54, 58, 513})  # токен, нет прав, незавершённая регистрация, логин не в Директе
_LIMIT_ERROR_CODES = frozenset({56, 152, 506, 9000})  # запросы, баллы, соединения, очередь отчётов
_TRANSIENT_STATUSES = frozenset({500, 502, 503, 504})
_GOAL_COLUMN = re.compile(r"^Conversions_(?P<goal>\d+)_(?P<model>[A-Z]+)$")


class DirectApiError(RuntimeError):
    """Ошибка Reports API: в тексте — сообщение Директа, HTTP-статус, код ошибки и RequestId."""

    def __init__(
        self, message: str, *, status: int | None = None, code: int | None = None, request_id: str | None = None
    ) -> None:
        super().__init__(message)
        self.status = status
        self.code = code
        self.request_id = request_id

    def __str__(self) -> str:
        details = [
            f"HTTP {self.status}" if self.status else "",
            f"код {self.code}" if self.code else "",
            f"RequestId {self.request_id}" if self.request_id else "",
        ]
        extra = ", ".join(detail for detail in details if detail)
        return f"{self.args[0]} ({extra})" if extra else self.args[0]


class DirectAuthError(DirectApiError):
    """Токена нет, он недействителен или у него нет доступа к кабинету клиента."""


class DirectLimitError(DirectApiError):
    """Исчерпаны баллы API, лимит запросов или очередь отчётов."""


class DirectReportTimeout(DirectApiError):
    """Отчёт не сформировался за отведённое число попыток."""


@dataclass(frozen=True, slots=True)
class ReportRequest:
    date_from: dt.date
    date_to: dt.date
    body: dict[str, Any]

    @property
    def name(self) -> str:
        return self.body["params"]["ReportName"]


def build_report_request(config: ReportConfig, today: dt.date | None = None) -> ReportRequest:
    """Тело запроса CUSTOM_DATE за период отчёта по правилам проекта (по умолчанию — до вчера).

    Поле Conversions запрашивается всегда: без целей это конверсии по всем целям,
    а с параметром Goals Директ разбивает его на колонки Conversions_<цель>_<модель>.
    """
    date_from, date_to = config.report_metadata.resolve_period(today=today)
    direct = config.direct_api
    params: dict[str, Any] = {
        "SelectionCriteria": {"DateFrom": date_from.isoformat(), "DateTo": date_to.isoformat()},
        "FieldNames": [*BASE_FIELDS, "Conversions"],
        "ReportType": REPORT_TYPE,
        "DateRangeType": "CUSTOM_DATE",
        "Format": "TSV",
        "IncludeVAT": "YES" if direct.include_vat else "NO",
    }
    if direct.goals:
        params["Goals"] = [str(goal) for goal in direct.goals]
    params["ReportName"] = _report_name(params, date_from, date_to)
    return ReportRequest(date_from=date_from, date_to=date_to, body={"params": params})


def _report_name(params: dict[str, Any], date_from: dt.date, date_to: dt.date) -> str:
    """Имя однозначно задаётся параметрами: повтор того же запроса опрашивает статус
    уже поставленного отчёта, а не ставит в очередь новый (имена в очереди уникальны)."""
    digest = hashlib.blake2b(json.dumps(params, sort_keys=True).encode("utf-8"), digest_size=6).hexdigest()
    return f"report-generator {date_from:%Y-%m-%d}..{date_to:%Y-%m-%d} {digest}"


def build_headers(token: str, config: DirectApiConfig) -> dict[str, str]:
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept-Language": "ru",  # тексты ошибок Директа — на русском
        "returnMoneyInMicros": "false",  # суммы в валюте, а не в миллионных долях
        "skipReportHeader": "true",
        "skipReportSummary": "true",
        "processingMode": "auto",
    }
    if config.client_login:
        headers["Client-Login"] = config.client_login
    return headers


def read_env_file(path: Path) -> dict[str, str]:
    """Разбор .env: строки KEY=VALUE, комментарии «#», значения в кавычках."""
    if not path.is_file():
        return {}
    values: dict[str, str] = {}
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip().removeprefix("export ").strip()
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        values[key] = value
    return values


def resolve_token(token: str | None = None) -> str:
    """OAuth-токен: аргумент → переменная окружения YANDEX_DIRECT_TOKEN → файл .env проекта."""
    candidates = (token, os.environ.get(TOKEN_ENV_VAR), read_env_file(ENV_FILE).get(TOKEN_ENV_VAR))
    for candidate in candidates:
        if candidate and candidate.strip():
            return candidate.strip()
    raise DirectAuthError(
        f"Не задан OAuth-токен Директа: задайте переменную окружения {TOKEN_ENV_VAR} "
        f"или добавьте её в файл {ENV_FILE}"
    )


def fetch_campaign_report(
    config: ReportConfig,
    *,
    token: str | None = None,
    today: dt.date | None = None,
    client: httpx.Client | None = None,
    url: str = REPORTS_URL,
    max_attempts: int = MAX_ATTEMPTS,
    sleep: Callable[[float], None] = time.sleep,
) -> list[DailyAdRecord]:
    """Скачивает суточную статистику кампаний за период отчёта и валидирует её.

    Args:
        config: конфиг отчёта — период и секция direct_api.
        token: OAuth-токен; по умолчанию — из окружения или .env (``resolve_token``).
        today: «сегодня» для расчёта периода.
        client: готовый httpx-клиент (в тестах — с MockTransport).
        url: адрес сервиса; для песочницы — ``SANDBOX_REPORTS_URL``.
        max_attempts: сколько раз запрашивать отчёт, пока он в очереди.
        sleep: функция ожидания между попытками.

    Raises:
        DirectAuthError: нет токена, он недействителен или нет доступа к кабинету.
        DirectLimitError: исчерпаны баллы, лимит запросов или очередь отчётов.
        DirectReportTimeout: отчёт не готов после ``max_attempts`` попыток.
        DirectApiError: прочие ошибки API и сети.
        ValueError: строка отчёта не прошла валидацию DailyAdRecord.
    """
    resolved_token = resolve_token(token)
    request = build_report_request(config, today)
    direct = config.direct_api
    headers = build_headers(resolved_token, direct)
    logger.info(
        "Директ: запрос отчёта за %s – %s (логин: %s; цели: %s; расход %s)",
        f"{request.date_from:%d.%m.%Y}", f"{request.date_to:%d.%m.%Y}",
        direct.client_login or "владелец токена",
        ", ".join(map(str, direct.goals)) or "все",
        "с НДС" if direct.include_vat else "без НДС",
    )  # fmt: skip

    if client is None:
        with httpx.Client(timeout=_TIMEOUT) as own_client:
            text = download_report(own_client, url, request, headers, max_attempts=max_attempts, sleep=sleep)
    else:
        text = download_report(client, url, request, headers, max_attempts=max_attempts, sleep=sleep)

    records = parse_report_tsv(text, direct.goals)
    logger.info("Директ: скачано строк: %d", len(records))
    return records


def download_report(
    client: httpx.Client,
    url: str,
    request: ReportRequest,
    headers: dict[str, str],
    *,
    max_attempts: int = MAX_ATTEMPTS,
    sleep: Callable[[float], None] = time.sleep,
) -> str:
    """Отправляет запрос и ждёт готовности отчёта; возвращает TSV.

    HTTP 201/202 — отчёт в очереди или формируется: ждём ``retryIn`` секунд и повторяем.
    Сбои сети и HTTP 5xx тоже повторяются: по документации Директа это временные ошибки.
    """
    last_error: DirectApiError | None = None
    for attempt in range(1, max_attempts + 1):
        logger.debug("Директ: POST %s, отчёт «%s», попытка %d из %d", url, request.name, attempt, max_attempts)
        wait = DEFAULT_RETRY_IN
        try:
            response = client.post(url, json=request.body, headers=headers)
        except httpx.TransportError as exc:
            last_error = DirectApiError(f"Нет связи с API Директа: {exc}")
            logger.warning("Директ: %s — повтор через %s", last_error, _format_wait(wait))
        else:
            status = response.status_code
            if status == 200:
                logger.debug("Директ: отчёт получен, RequestId %s", response.headers.get("RequestId"))
                return response.content.decode("utf-8-sig")
            wait = _retry_in(response)
            if status in (201, 202):
                queue = response.headers.get("reportsInQueue")
                logger.info(
                    "Директ: отчёт %s (HTTP %d%s) — повтор через %s, попытка %d из %d",
                    "поставлен в очередь" if status == 201 else "формируется", status,
                    f", отчётов в очереди: {queue}" if queue else "", _format_wait(wait), attempt, max_attempts,
                )  # fmt: skip
                last_error = DirectReportTimeout(
                    f"Отчёт «{request.name}» не готов после {max_attempts} попыток — Директ продолжит "
                    f"формировать его, повторите запуск позже",
                    status=status,
                    request_id=response.headers.get("RequestId"),
                )
            elif status in _TRANSIENT_STATUSES:
                last_error = _error_from(response)
                logger.warning("Директ: временная ошибка сервера — %s; повтор через %s", last_error, _format_wait(wait))
            else:
                raise _error_from(response)
        if attempt < max_attempts:
            sleep(wait)
    assert last_error is not None  # цикл выполняется хотя бы раз: max_attempts ≥ 1
    raise last_error


def parse_report_tsv(text: str, goals: Sequence[int] = ()) -> list[DailyAdRecord]:
    """TSV отчёта → DailyAdRecord через общую валидацию выгрузок (``parse_ad_rows``).

    Директ отдаёт TSV без кавычек, «--» вместо пустых значений; первая строка —
    заголовки колонок. С целями конверсии суммируются по колонкам целей из конфига.
    """
    reader = csv.DictReader(io.StringIO(text), delimiter="\t", quoting=csv.QUOTE_NONE)
    rows: list[dict[str, Any]] = list(reader)
    if goals:
        rows = [_with_goal_conversions(row, goals) for row in rows]
    return parse_ad_rows(rows, SOURCE_NAME)


def _with_goal_conversions(row: dict[str, Any], goals: Sequence[int]) -> dict[str, Any]:
    """Колонки Conversions_<цель>_<модель> → одна колонка Conversions: сумма по целям из конфига."""
    per_goal: dict[int, str] = {}
    for column, value in row.items():
        match = _GOAL_COLUMN.match(column or "")
        if not match:
            continue
        goal = int(match["goal"])
        if goal in per_goal:
            raise DirectApiError(f"В отчёте несколько колонок конверсий для цели {goal}: ожидается одна модель атрибуции")
        per_goal[goal] = value
    missing = [str(goal) for goal in goals if goal not in per_goal]
    if missing:
        raise DirectApiError(f"В отчёте нет колонок конверсий для целей: {', '.join(missing)}")
    return {**row, "Conversions": sum(_parse_count(per_goal[goal]) for goal in goals)}


def _parse_count(value: str | None) -> int:
    cleaned = (value or "").strip()
    if cleaned in {"", "--"}:  # «--» — нет данных за день
        return 0
    try:
        return int(cleaned)
    except ValueError:
        raise DirectApiError(f"Некорректное число конверсий в отчёте: «{value}»") from None


def _retry_in(response: httpx.Response) -> float:
    raw = response.headers.get("retryIn")
    try:
        return max(float(raw), 0.0) if raw is not None else DEFAULT_RETRY_IN
    except ValueError:
        return DEFAULT_RETRY_IN


def _format_wait(seconds: float) -> str:
    return f"{seconds:g} с"


def _error_from(response: httpx.Response) -> DirectApiError:
    """Ответ с ошибкой → исключение нужного типа с текстом Директа."""
    status = response.status_code
    request_id = response.headers.get("RequestId")
    try:
        error = response.json().get("error") or {}
    except (ValueError, AttributeError):  # не JSON (например, HTML от балансировщика)
        error = {}
    code = _parse_error_code(error.get("error_code"))
    request_id = error.get("request_id") or request_id
    message = ". ".join(str(part) for part in (error.get("error_string"), error.get("error_detail")) if part)
    message = message or f"API Директа ответил HTTP {status}: {response.text.strip()[:200] or response.reason_phrase}"

    if status in (401, 403) or code in _AUTH_ERROR_CODES:
        return DirectAuthError(
            f"Ошибка авторизации в API Директа: {message}. Проверьте {TOKEN_ENV_VAR} и direct_api.client_login",
            status=status, code=code, request_id=request_id,
        )  # fmt: skip
    if status == 429 or code in _LIMIT_ERROR_CODES:
        return DirectLimitError(f"Лимит API Директа: {message}", status=status, code=code, request_id=request_id)
    return DirectApiError(f"Ошибка API Директа: {message}", status=status, code=code, request_id=request_id)


def _parse_error_code(raw: Any) -> int | None:
    try:
        return int(raw) if raw is not None else None
    except (TypeError, ValueError):
        return None
