"""Источники рекламных данных: выгрузки кабинета (CSV/TSV/JSON) и генератор
тестовой суточной статистики в формате Яндекс Директа.

Проверочный запуск из корня проекта:
    python -m src.data_loader
    python src/data_loader.py
"""

import csv
import datetime as dt
import hashlib
import io
import json
import logging
import random
import sys
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from pydantic import ValidationError

if not __package__:  # запуск файлом (python src/data_loader.py): делаем пакет src видимым
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.console import configure_logging, use_utf8_console  # noqa: E402
from src.schemas import (  # noqa: E402
    CampaignSummary,
    DailyAdRecord,
    PerformanceMetrics,
    aggregate_by_campaign,
    calculate_totals,
)

logger = logging.getLogger(__name__)


@dataclass(frozen=True, slots=True)
class _ChannelProfile:
    """Диапазоны базовых параметров кампании; конкретные значения выбираются случайно."""

    daily_impressions: tuple[int, int]
    ctr: tuple[float, float]  # доля: 0.025 = 2,5 %
    cpc: tuple[float, float]  # цена клика в валюте кабинета
    cr: tuple[float, float]  # доля кликов, ставших конверсией


# Поиск: мало показов, высокий CTR и дорогой клик; РСЯ — наоборот.
# Диапазоны заложены с запасом внутри целевых CTR 1,5–3,5 % и CR 2–5 %,
# чтобы дневной шум не выводил итог кампании за эти границы.
_CHANNEL_PROFILES: dict[str, _ChannelProfile] = {
    "search": _ChannelProfile(
        daily_impressions=(2_500, 6_000), ctr=(0.024, 0.032), cpc=(35.0, 65.0), cr=(0.030, 0.042)
    ),
    "network": _ChannelProfile(
        daily_impressions=(10_000, 24_000), ctr=(0.016, 0.022), cpc=(10.0, 22.0), cr=(0.024, 0.034)
    ),
}

# Нейминг кампаний — как в выгрузках кабинета клиента.
_CAMPAIGN_TEMPLATES: tuple[tuple[str, str], ...] = (
    ("Поиск_Москва_Категории_Целевой", "search"),
    ("Поиск_РФ_Бренд_Транзакционный", "search"),
    ("РСЯ_РФ_Интересы_Аудитории", "network"),
    ("РСЯ_РФ_Ретаргетинг_Корзины", "network"),
    ("Поиск_СПб_Конкуренты_Широкий", "search"),
    ("РСЯ_Москва_Похожие_Покупатели", "network"),
    ("Поиск_РФ_Общие_Информационный", "search"),
    ("РСЯ_РФ_Смарт-баннеры_Фид", "network"),
)

# Спрос по дням недели (пн … вс): в будни выше, чем в выходные.
_WEEKDAY_FACTORS = (1.04, 1.07, 1.05, 1.02, 0.97, 0.90, 0.95)

# Сигмы логнормальных множителей ≈ относительный разброс день ко дню.
_NOISE_IMPRESSIONS = 0.12
_NOISE_CTR = 0.10
_NOISE_CPC = 0.10
_NOISE_CR = 0.15
_MAX_TREND = 0.10  # плавный рост или спад объёма к концу периода, до ±10 %

_CAMPAIGN_ID_RANGE = range(70_000_000, 120_000_000)  # 8–9-значные ID, как в Директе


def _noise(rng: random.Random, sigma: float) -> float:
    """Положительный множитель со средним ровно 1: сдвиг mu убирает смещение exp(σ²/2)."""
    return rng.lognormvariate(-sigma**2 / 2, sigma)


@dataclass(frozen=True, slots=True)
class _MockCampaign:
    campaign_id: int
    name: str
    base_impressions: float
    ctr: float
    cpc: float
    cr: float
    trend: float


def _build_campaigns(count: int, rng: random.Random) -> list[_MockCampaign]:
    campaigns = []
    for index, campaign_id in enumerate(rng.sample(_CAMPAIGN_ID_RANGE, count)):
        name, channel = _CAMPAIGN_TEMPLATES[index % len(_CAMPAIGN_TEMPLATES)]
        cycle = index // len(_CAMPAIGN_TEMPLATES)
        if cycle:
            name = f"{name}_{cycle + 1}"
        profile = _CHANNEL_PROFILES[channel]
        campaigns.append(
            _MockCampaign(
                campaign_id=campaign_id,
                name=name,
                base_impressions=rng.uniform(*profile.daily_impressions),
                ctr=rng.uniform(*profile.ctr),
                cpc=rng.uniform(*profile.cpc),
                cr=rng.uniform(*profile.cr),
                trend=rng.uniform(-_MAX_TREND, _MAX_TREND),
            )
        )
    return campaigns


def generate_mock_ad_data(
    days: int = 30,
    campaigns_count: int = 3,
    *,
    end_date: dt.date | None = None,
    seed: int | None = None,
) -> list[DailyAdRecord]:
    """Генерирует суточную статистику кампаний, провалидированную моделью DailyAdRecord.

    Модель дня: показы = база × день недели × тренд × шум; клики ~ Binomial(показы, CTR дня);
    расход = клики × CPC дня; конверсии ~ Binomial(клики, CR дня). Биномиальная выборка
    даёт целые значения с естественным разбросом, включая дни без конверсий.

    Args:
        days: длина периода в днях, не меньше 1.
        campaigns_count: число кампаний, не меньше 1.
        end_date: последний день периода включительно; по умолчанию вчера —
            статистика за сегодня в кабинете ещё неполная.
        seed: зерно генератора для воспроизводимых данных.

    Returns:
        Записи по дням; внутри дня — в порядке кампаний.
    """
    if days < 1:
        raise ValueError(f"days должен быть не меньше 1, получено {days}")
    if campaigns_count < 1:
        raise ValueError(f"campaigns_count должен быть не меньше 1, получено {campaigns_count}")

    rng = random.Random(seed)
    last_day = end_date or dt.date.today() - dt.timedelta(days=1)
    first_day = last_day - dt.timedelta(days=days - 1)
    campaigns = _build_campaigns(campaigns_count, rng)

    records: list[DailyAdRecord] = []
    for offset in range(days):
        day = first_day + dt.timedelta(days=offset)
        progress = offset / (days - 1) if days > 1 else 0.0
        for campaign in campaigns:
            demand = _WEEKDAY_FACTORS[day.weekday()] * (1 + campaign.trend * progress)
            impressions = round(campaign.base_impressions * demand * _noise(rng, _NOISE_IMPRESSIONS))
            day_ctr = min(1.0, campaign.ctr * _noise(rng, _NOISE_CTR))
            clicks = rng.binomialvariate(impressions, day_ctr)
            cost = round(clicks * campaign.cpc * _noise(rng, _NOISE_CPC), 2)
            day_cr = min(1.0, campaign.cr * _noise(rng, _NOISE_CR))
            conversions = rng.binomialvariate(clicks, day_cr)
            records.append(
                DailyAdRecord(
                    date=day,
                    campaign_id=campaign.campaign_id,
                    campaign_name=campaign.name,
                    impressions=impressions,
                    clicks=clicks,
                    cost=cost,
                    conversions=conversions,
                )
            )
    logger.debug("Синтетические данные: записей %d (кампаний: %d, дней: %d)", len(records), campaigns_count, days)
    return records


# ---------------------------------------------------------------------------
# Выгрузки кабинета
# ---------------------------------------------------------------------------

SYNTHETIC_ID_BASE = 10**12  # синтетические ID 13-значные — не пересекаются с 8–9-значными ID Директа


def campaign_id_from_name(name: str) -> int:
    """Стабильный ID кампании для выгрузок без колонки CampaignId.

    Берём blake2b, а не hash(): встроенный хеш строк меняется между запусками
    интерпретатора, а ID должен совпадать в отчётах за разные периоды.
    """
    digest = hashlib.blake2b(name.strip().encode("utf-8"), digest_size=8).digest()
    return SYNTHETIC_ID_BASE + int.from_bytes(digest) % SYNTHETIC_ID_BASE


def _with_campaign_id(row: dict[str, Any]) -> dict[str, Any]:
    if row.get("CampaignId") not in (None, "") or row.get("campaign_id") not in (None, ""):
        return row
    name = row.get("CampaignName", row.get("campaign_name"))
    if not isinstance(name, str):
        return row  # отсутствие названия сообщит валидация DailyAdRecord
    return {**row, "CampaignId": campaign_id_from_name(name)}


def _read_rows(source: Path) -> list[Any]:
    suffix = source.suffix.lower()
    if suffix == ".json":
        rows = json.loads(source.read_text(encoding="utf-8-sig"))
        if not isinstance(rows, list):
            raise ValueError(f"{source.name}: ожидается JSON-массив записей")
        return rows
    if suffix in {".csv", ".tsv"}:
        text = source.read_text(encoding="utf-8-sig")
        if suffix == ".tsv":
            delimiter = "\t"
        else:
            try:  # Excel с русской локалью сохраняет CSV через «;»
                delimiter = csv.Sniffer().sniff("\n".join(text.splitlines()[:5]), ",;\t").delimiter
            except csv.Error:
                delimiter = ","
        return list(csv.DictReader(io.StringIO(text), delimiter=delimiter))
    raise ValueError(f"{source.name}: неподдерживаемый формат, ожидается .csv, .tsv или .json")


def load_ad_records(path: str | Path) -> list[DailyAdRecord]:
    """Читает выгрузку кабинета и валидирует каждую строку моделью DailyAdRecord.

    Колонки — в нотации Reports API Директа (Date, CampaignName, Cost…). Если
    CampaignId нет, он выводится из названия кампании (``campaign_id_from_name``).

    Raises:
        ValueError: неподдерживаемый формат или невалидная строка (с её номером).
    """
    source = Path(path)
    return parse_ad_rows(_read_rows(source), source.name)


def parse_ad_rows(rows: Iterable[Any], source_name: str) -> list[DailyAdRecord]:
    """Валидирует строки выгрузки (словари с колонками Директа) — общий путь для файлов и API.

    Raises:
        ValueError: строка не словарь или не прошла валидацию; в сообщении — её номер.
    """
    records = []
    synthetic_ids = 0
    for number, row in enumerate(rows, start=1):
        if not isinstance(row, dict):
            raise ValueError(f"{source_name}, запись {number}: ожидается объект, получено {row!r}")
        prepared = _with_campaign_id(row)
        synthetic_ids += prepared is not row
        try:
            records.append(DailyAdRecord.model_validate(prepared))
        except ValidationError as exc:
            raise ValueError(f"{source_name}, запись {number}: {exc}") from exc
    if synthetic_ids:
        logger.info("%s: нет CampaignId — ID %d записей выведены из названий кампаний", source_name, synthetic_ids)
    logger.debug("%s: прочитано и провалидировано %d записей", source_name, len(records))
    return records


# ---------------------------------------------------------------------------
# Консольная сводка для проверочного запуска
# ---------------------------------------------------------------------------

_MISSING = "—"  # метрика не определена: нулевой знаменатель


def _format_number(value: float | None, digits: int = 0, suffix: str = "") -> str:
    """Русский формат: пробел между разрядами, запятая перед дробной частью."""
    if value is None:
        return _MISSING
    return f"{value:,.{digits}f}".replace(",", " ").replace(".", ",") + suffix


def _metric_cells(metrics: PerformanceMetrics) -> list[str]:
    return [
        _format_number(metrics.impressions),
        _format_number(metrics.clicks),
        _format_number(metrics.ctr, 2, "%"),
        _format_number(metrics.cpc, 2),
        _format_number(metrics.cost, 2),
        _format_number(metrics.conversions),
        _format_number(metrics.cr, 2, "%"),
        _format_number(metrics.cpa, 2),
    ]


def format_summary_table(
    summaries: Sequence[CampaignSummary],
    totals: PerformanceMetrics,
    currency: str = "₽",
) -> str:
    """Текстовая таблица сводки по кампаниям со строкой «ИТОГО»."""
    headers = [
        "Кампания", "ID", "Дней", "Показы", "Клики", "CTR", f"CPC, {currency}",
        f"Расход, {currency}", "Конв.", "CR", f"CPA, {currency}",
    ]  # fmt: skip
    rows = [
        [summary.campaign_name, str(summary.campaign_id), str(summary.days_count),
         *_metric_cells(summary)]
        for summary in summaries
    ]  # fmt: skip
    total_row = ["ИТОГО", "", "", *_metric_cells(totals)]

    table = [headers, *rows, total_row]
    widths = [max(len(row[column]) for row in table) for column in range(len(headers))]

    def render(row: Sequence[str]) -> str:
        cells = [row[0].ljust(widths[0])]
        cells += [cell.rjust(width) for cell, width in zip(row[1:], widths[1:])]
        return " | ".join(cells)

    separator = "-+-".join("-" * width for width in widths)
    return "\n".join([render(headers), separator, *map(render, rows), separator, render(total_row)])


def main() -> None:
    """Проверочный прогон: генерация → валидация → агрегация → сводная таблица."""
    use_utf8_console()
    configure_logging()

    records = generate_mock_ad_data(days=30)
    summaries = aggregate_by_campaign(records)
    totals = calculate_totals(summaries)
    date_from = min(record.date for record in records)
    date_to = max(record.date for record in records)

    logger.info("DailyAdRecord: сгенерировано и провалидировано %d суточных записей", len(records))
    logger.info("Период: %s – %s, кампаний: %d", f"{date_from:%d.%m.%Y}", f"{date_to:%d.%m.%Y}", len(summaries))
    logger.info("Сводка по кампаниям:\n%s", format_summary_table(summaries, totals))


if __name__ == "__main__":
    main()
