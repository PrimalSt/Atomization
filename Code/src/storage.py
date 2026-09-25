"""Локальная история помесячных итогов (SQLite) — база для динамики месяц к месяцу (MoM).

База — один файл, по умолчанию ``local_storage/history.db`` рядом с приложением;
папка и таблица создаются при первой записи. Снимок месяца — итоги отчёта клиента
за этот месяц. Повторная сборка отчёта за тот же месяц перезаписывает снимок
(UPSERT), поэтому исправленная выгрузка просто заменяет старые цифры.

Каждая функция открывает своё соединение: модуль можно вызывать из рабочего
потока GUI, не заботясь о том, в каком потоке соединение создано.
"""

import datetime as dt
import logging
import re
import sqlite3
from collections.abc import Iterator, Mapping, Sequence
from contextlib import closing, contextmanager
from pathlib import Path
from typing import Any

from src.paths import STORAGE_DIR
from src.schemas import CampaignSummary, PerformanceMetrics, calculate_totals

DEFAULT_DB_PATH = STORAGE_DIR / "history.db"
SCHEMA_VERSION = 1

_MONTH_PATTERN = re.compile(r"(?P<year>\d{4})-(?P<month>0[1-9]|1[0-2])")
_BUSY_TIMEOUT_SECONDS = 10  # вторая копия приложения пишет в ту же базу — ждём, а не падаем

_CREATE_TABLE = """
CREATE TABLE IF NOT EXISTS monthly_snapshots (
    client_id         TEXT    NOT NULL CHECK (length(client_id) > 0),
    report_month      TEXT    NOT NULL CHECK (report_month GLOB '[0-9][0-9][0-9][0-9]-[0-1][0-9]'),
    total_spend       REAL    NOT NULL CHECK (total_spend >= 0),
    total_clicks      INTEGER NOT NULL CHECK (total_clicks >= 0),
    total_impressions INTEGER NOT NULL CHECK (total_impressions >= 0),
    total_conversions INTEGER NOT NULL CHECK (total_conversions >= 0),
    avg_cpa           REAL,
    avg_ctr           REAL,
    created_at        TEXT    NOT NULL,
    updated_at        TEXT    NOT NULL,
    PRIMARY KEY (client_id, report_month)
)
"""

_UPSERT = """
INSERT INTO monthly_snapshots (
    client_id, report_month, total_spend, total_clicks, total_impressions, total_conversions,
    avg_cpa, avg_ctr, created_at, updated_at
) VALUES (
    :client_id, :report_month, :total_spend, :total_clicks, :total_impressions, :total_conversions,
    :avg_cpa, :avg_ctr, :now, :now
)
ON CONFLICT (client_id, report_month) DO UPDATE SET
    total_spend = excluded.total_spend,
    total_clicks = excluded.total_clicks,
    total_impressions = excluded.total_impressions,
    total_conversions = excluded.total_conversions,
    avg_cpa = excluded.avg_cpa,
    avg_ctr = excluded.avg_ctr,
    updated_at = excluded.updated_at
"""

_SELECT = "SELECT * FROM monthly_snapshots WHERE client_id = ? AND report_month = ?"

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Месяцы
# ---------------------------------------------------------------------------


def month_key(day: dt.date) -> str:
    """Дата → месяц в формате хранилища: 2026-09-21 → «2026-09»."""
    return f"{day:%Y-%m}"


def parse_month(month: str) -> tuple[int, int]:
    """«2026-09» → (2026, 9).

    Raises:
        ValueError: строка не в формате YYYY-MM.
    """
    match = _MONTH_PATTERN.fullmatch(month.strip()) if isinstance(month, str) else None
    if match is None:
        raise ValueError(f"Месяц отчёта должен быть в формате YYYY-MM, получено «{month}»")
    return int(match["year"]), int(match["month"])


def previous_month(month: str) -> str:
    """«2026-01» → «2025-12»."""
    year, number = parse_month(month)
    return f"{year - 1}-12" if number == 1 else f"{year}-{number - 1:02d}"


def report_month_for(period: tuple[dt.date, dt.date]) -> str:
    """Месяц, к которому относится отчёт: тот, на который приходится больше дней периода.

    Выгрузка за календарный месяц даёт этот месяц; скользящие 30 дней с 23.08 по 21.09 —
    сентябрь. При равенстве берётся более поздний месяц.
    """
    date_from, date_to = period
    if date_from > date_to:
        raise ValueError(f"Начало периода ({date_from}) позже конца ({date_to})")
    days: dict[str, int] = {}
    day = date_from
    while day <= date_to:
        days[month_key(day)] = days.get(month_key(day), 0) + 1
        day += dt.timedelta(days=1)
    return max(days, key=lambda month: (days[month], month))


def is_full_month(period: tuple[dt.date, dt.date]) -> bool:
    """Период — ровно один календарный месяц, с первого по последнее число."""
    date_from, date_to = period
    next_day = date_to + dt.timedelta(days=1)
    return date_from.day == 1 and next_day.day == 1 and month_key(date_from) == month_key(date_to)


# ---------------------------------------------------------------------------
# База
# ---------------------------------------------------------------------------


def _check_client_id(client_id: str) -> str:
    if not isinstance(client_id, str) or not client_id.strip():
        raise ValueError("client_id не может быть пустым")
    return client_id.strip()


def _now() -> str:
    """Местное время с поясом: «2026-09-25T22:10:00+03:00»."""
    return dt.datetime.now().astimezone().isoformat(timespec="seconds")


@contextmanager
def _connect(db_path: str | Path) -> Iterator[sqlite3.Connection]:
    """Соединение с готовой схемой; блок внутри — одна транзакция (commit или rollback)."""
    path = Path(db_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with closing(sqlite3.connect(path, timeout=_BUSY_TIMEOUT_SECONDS)) as connection:
        connection.row_factory = sqlite3.Row
        with connection:
            _migrate(connection)
            yield connection


def _migrate(connection: sqlite3.Connection) -> None:
    """Создаёт схему в новой базе; номер версии — в PRAGMA user_version для будущих миграций."""
    version = connection.execute("PRAGMA user_version").fetchone()[0]
    if version > SCHEMA_VERSION:
        raise sqlite3.DatabaseError(
            f"База истории создана более новой версией приложения (схема {version}, поддерживается {SCHEMA_VERSION})"
        )
    connection.execute(_CREATE_TABLE)
    if version < SCHEMA_VERSION:
        connection.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")


def init_storage(db_path: str | Path = DEFAULT_DB_PATH) -> Path:
    """Создаёт файл базы и таблицу, если их нет. Возвращает абсолютный путь к базе."""
    with _connect(db_path):
        pass
    return Path(db_path).resolve()


def save_monthly_metrics(
    client_id: str,
    month: str,
    summary: Sequence[CampaignSummary],
    *,
    db_path: str | Path = DEFAULT_DB_PATH,
) -> dict[str, Any]:
    """Сохраняет итоги клиента за месяц или обновляет уже сохранённые (UPSERT).

    Итоги считаются от сумм по кампаниям, как строка «ИТОГО» отчёта: CPA = расход /
    конверсии, CTR = клики / показы × 100 (в процентах). Без конверсий или показов
    соответствующая средняя метрика сохраняется как NULL. При обновлении ``created_at``
    остаётся от первой записи, ``updated_at`` — время последней.

    Args:
        client_id: идентификатор клиента (генератор отчётов передаёт client_name из конфига).
        month: месяц отчёта, YYYY-MM.
        summary: сводки по кампаниям за месяц.
        db_path: файл базы SQLite.

    Returns:
        Сохранённый снимок — как его вернёт ``get_monthly_metrics``.

    Raises:
        ValueError: пустой client_id, месяц не в формате YYYY-MM или нет сводок.
        sqlite3.Error: база недоступна или повреждена.
    """
    client = _check_client_id(client_id)
    year, number = parse_month(month)
    if not summary:
        raise ValueError(f"Нет сводок по кампаниям за {month} — сохранять нечего")
    totals = calculate_totals(summary)
    values = {
        "client_id": client,
        "report_month": f"{year}-{number:02d}",
        "total_spend": totals.cost,
        "total_clicks": totals.clicks,
        "total_impressions": totals.impressions,
        "total_conversions": totals.conversions,
        "avg_cpa": totals.cpa,
        "avg_ctr": totals.ctr,
        "now": _now(),
    }
    with _connect(db_path) as connection:
        connection.execute(_UPSERT, values)
        saved = connection.execute(_SELECT, (client, values["report_month"])).fetchone()
    logger.debug("История: сохранены итоги «%s» за %s в %s", client, values["report_month"], db_path)
    return dict(saved)


def get_monthly_metrics(
    client_id: str,
    month: str,
    *,
    db_path: str | Path = DEFAULT_DB_PATH,
) -> dict[str, Any] | None:
    """Снимок клиента за месяц или None, если его нет (или базы ещё нет).

    Ключи словаря — колонки таблицы monthly_snapshots: client_id, report_month,
    total_spend, total_clicks, total_impressions, total_conversions, avg_cpa, avg_ctr,
    created_at, updated_at.
    """
    client = _check_client_id(client_id)
    year, number = parse_month(month)
    if not Path(db_path).is_file():
        return None  # чтение не создаёт пустую базу
    with _connect(db_path) as connection:
        row = connection.execute(_SELECT, (client, f"{year}-{number:02d}")).fetchone()
    return dict(row) if row is not None else None


def get_previous_month_metrics(
    client_id: str,
    current_month: str,
    *,
    db_path: str | Path = DEFAULT_DB_PATH,
) -> dict[str, Any] | None:
    """Снимок за календарный месяц перед ``current_month`` — база для динамики MoM.

    Берётся именно предыдущий месяц: если за него отчёт не собирали, возвращается None,
    а не более ранний месяц — сравнение через пропуск вводило бы в заблуждение.
    """
    return get_monthly_metrics(client_id, previous_month(current_month), db_path=db_path)


def snapshot_metrics(snapshot: Mapping[str, Any]) -> PerformanceMetrics:
    """Счётчики снимка → PerformanceMetrics: CPC, CR и остальные KPI считаются теми же формулами."""
    return PerformanceMetrics(
        impressions=snapshot["total_impressions"],
        clicks=snapshot["total_clicks"],
        cost=snapshot["total_spend"],
        conversions=snapshot["total_conversions"],
    )
