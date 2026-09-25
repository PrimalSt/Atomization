"""Источники рекламных данных: выгрузки кабинета (CSV/TSV/JSON, Excel .xlsx/.xls)
и генератор тестовой суточной статистики в формате Яндекс Директа.

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
from collections import Counter
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, get_args

import openpyxl
import xlrd
from pydantic import ValidationError

if not __package__:  # запуск файлом (python src/data_loader.py): делаем пакет src видимым
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.console import configure_logging, use_utf8_console  # noqa: E402
from src.schemas import (  # noqa: E402
    CampaignSummary,
    DailyAdRecord,
    InputColumn,
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

EXCEL_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xls"})
SUPPORTED_SUFFIXES = frozenset({".csv", ".tsv", ".json"}) | EXCEL_SUFFIXES

INPUT_COLUMNS: tuple[str, ...] = get_args(InputColumn)
# CampaignId необязателен: без него ID выводится из названия кампании.
REQUIRED_COLUMNS: tuple[str, ...] = tuple(column for column in INPUT_COLUMNS if column != "CampaignId")
# Те же поля под именами Python (snake_case) — так их принимает DailyAdRecord.
_FIELD_NAMES: dict[str, str] = {
    field.alias: name for name, field in DailyAdRecord.model_fields.items() if field.alias in INPUT_COLUMNS
}

_ZIP_SIGNATURE = b"PK\x03\x04"  # .xlsx / .xlsm — ZIP-архив Office Open XML
_OLE2_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # .xls — Excel 97–2003
_HEADER_SCAN_ROWS = 30  # над таблицей в выгрузках бывают заголовок отчёта, период, фильтры
_HEADER_MIN_KNOWN = 3  # столько известных колонок должно быть в строке, чтобы счесть её заголовком


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


# --- Маппинг колонок --------------------------------------------------------


def column_key(name: object) -> str:
    """Ключ сравнения названий колонок: «  Затраты,  руб. » и «затраты, руб.» совпадают."""
    return " ".join(str(name).split()).casefold()


def _mapping_lookup(column_mapping: Mapping[str, str] | None) -> dict[str, str]:
    return {column_key(source): target for source, target in (column_mapping or {}).items()}


def _rename_columns(row: Mapping[Any, Any], lookup: Mapping[str, str]) -> dict[Any, Any]:
    if not lookup:
        return dict(row)
    renamed: dict[Any, Any] = {}
    origins: dict[Any, Any] = {}
    for column, value in row.items():
        name = lookup.get(column_key(column), column) if isinstance(column, str) else column
        if name in renamed:
            raise ValueError(
                f"колонки «{origins[name]}» и «{column}» обе соответствуют {name} — уточните column_mapping конфига"
            )
        renamed[name] = value
        origins[name] = column
    return renamed


def normalize_columns(row: Mapping[str, Any], column_mapping: Mapping[str, str] | None) -> dict[str, Any]:
    """Переименовывает колонки строки выгрузки в нотацию Директа по ``column_mapping``.

    Названия сравниваются без учёта регистра и лишних пробелов: при маппинге
    ``{"Затраты": "Cost"}`` колонка «затраты » тоже станет Cost. Колонки, которых
    нет в маппинге, остаются как есть — в том числе уже названные по-директовски.

    Raises:
        ValueError: две колонки строки получают одно имя (например, в файле есть и
            «Cost», и «Затраты», а маппинг переименовывает «Затраты» в Cost).
    """
    return _rename_columns(row, _mapping_lookup(column_mapping))


def _check_required_columns(columns: Iterable[Any], source_name: str) -> None:
    present = {column for column in columns if isinstance(column, str)}
    missing = [column for column in REQUIRED_COLUMNS if column not in present and _FIELD_NAMES[column] not in present]
    if missing:
        found = ", ".join(f"«{column}»" for column in present) or "нет"
        raise ValueError(
            f"{source_name}: нет обязательных колонок {', '.join(missing)} (колонки файла: {found}). "
            "Если в выгрузке они называются иначе, сопоставьте их в секции column_mapping конфига"
        )


# --- Чтение файлов ----------------------------------------------------------


def _read_rows(source: Path, lookup: Mapping[str, str]) -> list[Any]:
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
    if suffix in EXCEL_SUFFIXES:
        return _read_excel(source, lookup)
    raise ValueError(f"{source.name}: неподдерживаемый формат, ожидается .csv, .tsv, .json, .xlsx или .xls")


SheetRows = list[tuple[Any, ...]]


def _read_excel(source: Path, lookup: Mapping[str, str]) -> list[dict[str, Any]]:
    """Строки таблицы из книги Excel: первый лист, где нашлась строка заголовков.

    Формат определяется по содержимому, а не по расширению: .xlsx, переименованный
    в .xls (и наоборот), тоже читается.
    """
    with source.open("rb") as book:
        signature = book.read(len(_OLE2_SIGNATURE))
    if signature.startswith(_ZIP_SIGNATURE):
        sheets = _xlsx_sheets(source)
    elif signature == _OLE2_SIGNATURE:
        sheets = _xls_sheets(source)
    else:
        raise ValueError(
            f"{source.name}: файл не похож на книгу Excel (.xlsx или .xls) — возможно, это CSV или HTML "
            "с другим расширением; пересохраните его в Excel как «Книга Excel (.xlsx)»"
        )
    return _table_rows(sheets, source.name, lookup)


def _xlsx_sheets(source: Path) -> list[tuple[str, SheetRows]]:
    # openpyxl получает поток, а не путь: по пути он проверяет расширение и отвергает
    # .xlsx, сохранённый как .xls, а формат здесь уже определён по содержимому.
    with source.open("rb") as stream:
        try:
            workbook = openpyxl.load_workbook(stream, read_only=True, data_only=True)
            try:
                return [
                    (sheet.title, [row for row in sheet.iter_rows(values_only=True) if not _is_blank_row(row)])
                    for sheet in workbook.worksheets
                ]
            finally:
                workbook.close()
        except OSError:
            raise
        except Exception as exc:  # битый архив или XML: у openpyxl нет общего класса ошибок
            raise ValueError(f"{source.name}: не удалось прочитать книгу Excel: {exc}") from exc


def _xls_sheets(source: Path) -> list[tuple[str, SheetRows]]:
    try:
        workbook = xlrd.open_workbook(source, on_demand=True)
    except xlrd.XLRDError as exc:
        raise ValueError(f"{source.name}: не удалось прочитать книгу Excel 97–2003: {exc}") from exc
    try:
        sheets = []
        for index in range(workbook.nsheets):
            sheet = workbook.sheet_by_index(index)
            rows = [tuple(_xls_value(cell, workbook.datemode) for cell in sheet.row(n)) for n in range(sheet.nrows)]
            sheets.append((sheet.name, [row for row in rows if not _is_blank_row(row)]))
        return sheets
    finally:
        workbook.release_resources()


def _xls_value(cell: xlrd.sheet.Cell, datemode: int) -> Any:
    match cell.ctype:
        case xlrd.XL_CELL_EMPTY | xlrd.XL_CELL_BLANK:
            return None
        case xlrd.XL_CELL_DATE:
            return xlrd.xldate.xldate_as_datetime(cell.value, datemode)
        case xlrd.XL_CELL_BOOLEAN:
            return bool(cell.value)
        case xlrd.XL_CELL_ERROR:
            return xlrd.error_text_from_code.get(cell.value, "#ERROR")
    return cell.value


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _is_blank_row(row: Sequence[Any]) -> bool:
    """Пустые строки отбрасываются ещё при чтении: оформленные, но пустые строки до конца листа не копятся."""
    return all(map(_is_blank, row))


def _cell_value(value: Any) -> Any:
    """Даты Excel приходят как datetime с нулевым временем — для суточной статистики это дата."""
    if isinstance(value, dt.datetime) and value.time() == dt.time():
        return value.date()
    return value


def _find_header(rows: SheetRows, lookup: Mapping[str, str]) -> int | None:
    for index, row in enumerate(rows[:_HEADER_SCAN_ROWS]):
        names = {lookup.get(column_key(cell), str(cell).strip()) for cell in row if not _is_blank(cell)}
        if len(names & set(INPUT_COLUMNS)) >= _HEADER_MIN_KNOWN:
            return index
    return None


def _locate_table(
    sheets: list[tuple[str, SheetRows]], lookup: Mapping[str, str]
) -> tuple[str, SheetRows, int] | None:
    """Лист и номер строки заголовков: (название листа, строки листа, индекс заголовка).

    Заголовок — первая строка (в пределах первых ``_HEADER_SCAN_ROWS``), где с учётом
    маппинга есть хотя бы ``_HEADER_MIN_KNOWN`` колонок Директа. Если такой нет ни на
    одном листе, заголовком считается первая непустая строка первого непустого листа:
    тогда проверка колонок перечислит, каких не хватает. None — книга пустая.
    """
    for title, rows in sheets:
        header_index = _find_header(rows, lookup)
        if header_index is not None:
            return title, rows, header_index
    for title, rows in sheets:
        if rows:  # пустые строки отброшены при чтении — первая строка и есть первая непустая
            return title, rows, 0
    return None


def _table_rows(
    sheets: list[tuple[str, SheetRows]], source_name: str, lookup: Mapping[str, str]
) -> list[dict[str, Any]]:
    """Строки таблицы под заголовком — словари «колонка → значение»; пустые строки пропускаются."""
    located = _locate_table(sheets, lookup)
    if located is None:
        return []
    title, rows, header_index = located

    header = ["" if _is_blank(cell) else " ".join(str(cell).split()) for cell in rows[header_index]]
    duplicates = sorted(name for name, count in Counter(filter(None, header)).items() if count > 1)
    if duplicates:
        raise ValueError(f"{source_name}, лист «{title}»: повторяются колонки {', '.join(duplicates)}")
    logger.debug("%s: лист «%s», колонки таблицы: %s", source_name, title, header)

    table = []
    for row in rows[header_index + 1 :]:
        cells = (list(row) + [None] * len(header))[: len(header)]  # ячейки правее заголовка не нужны
        table.append({name: _cell_value(value) for name, value in zip(header, cells, strict=True) if name})
    return table


def load_ad_records(path: str | Path, column_mapping: Mapping[str, str] | None = None) -> list[DailyAdRecord]:
    """Читает выгрузку кабинета и валидирует каждую строку моделью DailyAdRecord.

    Форматы: .csv (разделитель «,», «;» или табуляция), .tsv, .json (массив объектов),
    .xlsx / .xlsm и .xls (Excel 97–2003). Колонки — в нотации Reports API Директа
    (Date, CampaignName, Cost…); если в файле они называются иначе, ``column_mapping``
    переименовывает их до валидации (``{"Затраты": "Cost"}``). Если CampaignId нет,
    он выводится из названия кампании (``campaign_id_from_name``).

    Raises:
        ValueError: неподдерживаемый или повреждённый файл, нет обязательных колонок,
            невалидная строка (с её номером).
    """
    source = Path(path)
    return parse_ad_rows(_read_rows(source, _mapping_lookup(column_mapping)), source.name, column_mapping)


def parse_ad_rows(
    rows: Iterable[Any], source_name: str, column_mapping: Mapping[str, str] | None = None
) -> list[DailyAdRecord]:
    """Валидирует строки выгрузки (словари с колонками Директа) — общий путь для файлов и API.

    Колонки сначала переименовываются по ``column_mapping`` (см. ``normalize_columns``);
    по первой строке проверяется, что обязательные колонки есть.

    Raises:
        ValueError: строка не словарь, нет обязательных колонок или строка не прошла
            валидацию; в сообщении — номер строки.
    """
    lookup = _mapping_lookup(column_mapping)
    records = []
    synthetic_ids = 0
    for number, row in enumerate(rows, start=1):
        if not isinstance(row, dict):
            raise ValueError(f"{source_name}, запись {number}: ожидается объект, получено {row!r}")
        try:
            renamed = _rename_columns(row, lookup)
        except ValueError as exc:
            raise ValueError(f"{source_name}, запись {number}: {exc}") from exc
        if number == 1:
            _check_required_columns(renamed, source_name)
        prepared = _with_campaign_id(renamed)
        synthetic_ids += prepared is not renamed
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
