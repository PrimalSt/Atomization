"""Единый пайплайн отчёта: конфиг → данные → агрегация → презентация."""

import datetime as dt
import logging
import re
import time
from collections.abc import Iterator, Sequence
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pydantic import ValidationError

from src.config_loader import DEFAULT_CONFIG_PATH, load_report_config
from src.data_loader import generate_mock_ad_data, load_ad_records
from src.direct_client import fetch_campaign_report
from src.rendering.builder import OUTPUT_DIR, build_presentation
from src.rendering.formatting import format_value
from src.schemas import ExpertNotes, ReportConfig, aggregate_by_campaign, calculate_totals

logger = logging.getLogger(__name__)

DEFAULT_MOCK_CAMPAIGNS = 4
NOTES_SUFFIXES = frozenset({".txt", ".md", ".markdown"})
_STEPS_COUNT = 4

NotesSource = str | Sequence[str] | Path


@dataclass(frozen=True, slots=True)
class StepTiming:
    name: str
    seconds: float


@dataclass(frozen=True, slots=True)
class ReportResult:
    """Итог прогона: где лежит отчёт, что в него вошло и сколько заняли шаги."""

    output_path: Path
    period: tuple[dt.date, dt.date]
    records_count: int
    campaigns_count: int
    slides_count: int
    has_notes: bool
    timings: tuple[StepTiming, ...]

    @property
    def total_seconds(self) -> float:
        return sum(step.seconds for step in self.timings)

    def timing_summary(self) -> str:
        """«Конфиг 0,01 с · Данные 0,03 с · … · всего 0,25 с»."""
        steps = [f"{step.name} {_format_seconds(step.seconds)}" for step in self.timings]
        return " · ".join([*steps, f"всего {_format_seconds(self.total_seconds)}"])


def generate_report(
    config_path: str | Path = DEFAULT_CONFIG_PATH,
    *,
    input_path: str | Path | None = None,
    mock: bool = False,
    direct: bool = False,
    notes: NotesSource | None = None,
    output: str | Path = OUTPUT_DIR,
    mock_campaigns: int = DEFAULT_MOCK_CAMPAIGNS,
    mock_seed: int | None = None,
    direct_token: str | None = None,
    today: dt.date | None = None,
) -> ReportResult:
    """Выполняет полный цикл и сохраняет презентацию.

    Источник данных — ровно один из ``input_path``, ``mock`` и ``direct``.

    Args:
        config_path: YAML-конфиг отчёта.
        input_path: выгрузка кабинета (.csv, .tsv, .json).
        mock: сгенерировать синтетические данные за период из конфига.
        direct: скачать статистику из Reports API Директа за период из конфига.
        notes: выводы специалиста — текст (упрощённый Markdown), список пунктов или
            ``Path`` к файлу .txt / .md. Строка всегда считается текстом, а не путём.
        output: папка (имя файла сформируется из клиента и периода) или путь к .pptx.
        mock_campaigns: число кампаний в синтетических данных.
        mock_seed: зерно генератора для воспроизводимых синтетических данных.
        direct_token: OAuth-токен Директа; по умолчанию — YANDEX_DIRECT_TOKEN из окружения или .env.
        today: «сегодня» для периода отчёта и даты формирования.

    Raises:
        ValueError: источник данных не указан или указано несколько; конфиг, данные
            или заметки не прошли валидацию.
        FileNotFoundError: нет конфига, файла данных или файла заметок.
        PermissionError: файл отчёта открыт в другой программе.
        DirectApiError: ошибка API Директа (авторизация, лимиты, очередь отчётов).
    """
    sources = (input_path is not None) + mock + direct
    if sources > 1:
        raise ValueError("Укажите один источник данных: файл выгрузки, mock или API Директа")
    if sources == 0:
        raise ValueError("Не указан источник данных: передайте файл выгрузки, включите mock или direct")
    today = today or dt.date.today()
    timings: list[StepTiming] = []

    with _step(timings, 1, "Конфиг") as details:
        config = _load_config(config_path)
        expert_notes = _load_notes(notes)
        details.append(f"клиент «{config.report_metadata.client_name}», активных слайдов: {len(config.active_slides)}")
        details.append(f"выводы: {_describe_notes(expert_notes)}")
    _warn_if_notes_unused(config, expert_notes)

    with _step(timings, 2, "Данные") as details:
        if input_path is not None:
            source = _existing_file(input_path, "Файл данных")
            records = load_ad_records(source)
            if not records:
                raise ValueError(f"{source.name}: в выгрузке нет ни одной записи")
            details.append(f"{source.name}")
        elif direct:
            records = fetch_campaign_report(config, token=direct_token, today=today)
            if not records:
                raise ValueError(
                    "API Директа не вернул статистику за период — проверьте direct_api.client_login "
                    "и что кампании показывались"
                )
            details.append("API Директа")
        else:
            date_from, date_to = config.report_metadata.resolve_period(today=today)
            records = generate_mock_ad_data(
                days=(date_to - date_from).days + 1, campaigns_count=mock_campaigns, end_date=date_to, seed=mock_seed
            )
            details.append("синтетические")
        period = (min(record.date for record in records), max(record.date for record in records))
        details.append(f"записей: {len(records)}, период {_format_period(period)}")
    if not mock:
        _warn_about_period(config, period, today)

    with _step(timings, 3, "Агрегация") as details:
        summaries = aggregate_by_campaign(records)
        totals = calculate_totals(summaries)
        spend = format_value(totals.cost, "currency", config.report_metadata.currency)
        details.append(f"кампаний: {len(summaries)}, расход {spend}, конверсий: {totals.conversions}")

    with _step(timings, 4, "Рендеринг") as details:
        target = resolve_output_path(output, config.report_metadata.client_name, period)
        path = build_presentation(config, summaries, records, target, generated_at=today, notes=expert_notes)
        # Открываем сохранённый файл: это проверка, что он читается, и фактическое число
        # слайдов — длинные выводы специалиста занимают больше одного.
        slides_count = len(Presentation(path).slides)
        details.append(f"слайдов: {slides_count} → {path}")

    return ReportResult(
        output_path=path,
        period=period,
        records_count=len(records),
        campaigns_count=len(summaries),
        slides_count=slides_count,
        has_notes=expert_notes is not None,
        timings=tuple(timings),
    )


def resolve_output_path(output: str | Path, client_name: str, period: tuple[dt.date, dt.date]) -> Path:
    """Путь к .pptx как есть; для папки — имя из клиента и периода.

    ``output/`` → ``output/report_ООО_Доставка_Плюс_2026-08-23_2026-09-21.pptx``.
    """
    target = Path(output)
    if target.suffix.lower() == ".pptx":
        return target
    if target.suffix and not target.is_dir():
        raise ValueError(f"Ожидается папка или файл .pptx, получено «{target}»")
    slug = re.sub(r"\W+", "_", client_name).strip("_") or "client"
    date_from, date_to = period
    return target / f"report_{slug}_{date_from:%Y-%m-%d}_{date_to:%Y-%m-%d}.pptx"


@contextmanager
def _step(timings: list[StepTiming], number: int, name: str) -> Iterator[list[str]]:
    """Замеряет шаг и пишет в лог его итог; подробности шаг дописывает в отданный список."""
    details: list[str] = []
    logger.debug("[%d/%d] %s…", number, _STEPS_COUNT, name)
    started = time.perf_counter()
    yield details
    elapsed = time.perf_counter() - started
    timings.append(StepTiming(name, elapsed))
    logger.info("[%d/%d] %s: %s (%s)", number, _STEPS_COUNT, name, "; ".join(details), _format_seconds(elapsed))


def _existing_file(path: str | Path, what: str) -> Path:
    candidate = Path(path)
    if not candidate.is_file():
        raise FileNotFoundError(f"{what} не найден: {candidate}")
    return candidate


def _load_config(path: str | Path) -> ReportConfig:
    source = _existing_file(path, "Конфиг")
    try:
        return load_report_config(source)
    except ValidationError as exc:
        raise ValueError(f"Конфиг {source} не прошёл валидацию:\n{exc}") from exc


def _load_notes(notes: NotesSource | None) -> ExpertNotes | None:
    if notes is None:
        return None
    if isinstance(notes, Path):
        source = _existing_file(notes, "Файл заметок")
        if source.suffix.lower() not in NOTES_SUFFIXES:
            raise ValueError(f"Файл заметок должен быть .txt или .md, получено «{source.name}»")
        notes = _read_text(source)
    try:
        parsed = ExpertNotes.parse(notes)
    except ValidationError as exc:
        raise ValueError(f"Выводы специалиста не прошли валидацию:\n{exc}") from exc
    if parsed is None:
        logger.warning("Выводы специалиста пустые — на слайде останется заготовка")
    return parsed


def _read_text(source: Path) -> str:
    """UTF-8 (с BOM или без); старый «Блокнот» сохранял в cp1251 — читаем и его."""
    try:
        return source.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        logger.warning("%s: файл не в UTF-8, прочитан в кодировке cp1251", source.name)
        return source.read_text(encoding="cp1251")


def _describe_notes(notes: ExpertNotes | None) -> str:
    if notes is None:
        return "заготовка"
    items = sum(block.kind in {"bullet", "numbered"} for block in notes.blocks)
    return f"блоков: {len(notes.blocks)}, из них пунктов списка: {items}"


def _warn_if_notes_unused(config: ReportConfig, notes: ExpertNotes | None) -> None:
    if notes is not None and not any(slide.type == "notes_slide" for slide in config.active_slides):
        logger.warning("Выводы переданы, но в конфиге нет активного слайда notes_slide — они не попадут в отчёт")


def _warn_about_period(config: ReportConfig, period: tuple[dt.date, dt.date], today: dt.date) -> None:
    """Сверяет период выгрузки с конфигом и предупреждает о неполных данных."""
    date_from, date_to = period
    if date_to >= today:
        logger.warning(
            "Последний день выгрузки — %s: статистика за сегодняшний день ещё неполная", f"{date_to:%d.%m.%Y}"
        )
    span = (date_to - date_from).days + 1
    expected = config.report_metadata.period_days
    if span != expected:
        logger.warning(
            "Выгрузка охватывает %d дн. (%s), а в конфиге period_days: %d", span, _format_period(period), expected
        )


def _format_period(period: tuple[dt.date, dt.date]) -> str:
    date_from, date_to = period
    return f"{date_from:%d.%m.%Y} – {date_to:%d.%m.%Y}"


def _format_seconds(seconds: float) -> str:
    """«45 мс» для быстрых шагов, «1,27 с» — для долгих."""
    if seconds < 1:
        return f"{seconds * 1000:.0f} мс"
    return f"{seconds:.2f}".replace(".", ",") + " с"
