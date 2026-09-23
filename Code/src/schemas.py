"""Схемы слоя валидации: сырые данные кабинета, агрегаты по кампаниям и конфиг отчёта.

Поля данных в Python называются в snake_case, а их алиасы — в нотации Reports API
Яндекс Директа (Date, CampaignId, CampaignName, Impressions, Clicks, Cost, Conversions).
Поэтому модели принимают выгрузку кабинета без переименований, а
``model_dump(by_alias=True)`` отдаёт ровно те ключи, на которые ссылается
config/report_config.yaml (``field: "CPA"``, ``series: ["Cost", "Conversions"]``).
"""

import datetime as dt
import math
import re
from collections import Counter, defaultdict
from collections.abc import Iterable, Sequence
from typing import Annotated, Any, Literal, Self

from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    NonNegativeInt,
    PositiveInt,
    computed_field,
    field_validator,
    model_validator,
)
from pydantic.alias_generators import to_pascal

# ---------------------------------------------------------------------------
# Словари допустимых имён — общие для моделей данных и конфига отчёта
# ---------------------------------------------------------------------------

CounterField = Literal["Impressions", "Clicks", "Cost", "Conversions"]
RatioField = Literal["CTR", "CPC", "CPA", "CR"]
MetricField = Literal[CounterField, RatioField]
SummaryField = Literal["CampaignId", "CampaignName", "DaysCount", MetricField]

# Итоговые KPI по всем кампаниям (слайд kpi_cards).
KpiMetricId = Literal[
    "total_spend",
    "impressions",
    "clicks",
    "conversions",
    "avg_ctr",
    "avg_cpc",
    "avg_cpa",
    "avg_cr",
]
MetricFormat = Literal["currency", "integer", "percent", "decimal"]

# Деньги: неотрицательные и конечные — NaN/inf из битой выгрузки отсекаются.
Money = Annotated[float, Field(ge=0, allow_inf_nan=False)]

_HEX_COLOR_PATTERN = r"^[0-9A-F]{6}$"


def safe_divide(numerator: float, denominator: float, *, scale: float = 1.0) -> float | None:
    """Делит без ZeroDivisionError: при нулевом знаменателе метрика не определена.

    Возвращается None, а не 0 и не inf: CPA = 0 читался бы как «бесплатные лиды»,
    а inf ломает JSON-сериализацию и оси графиков.
    """
    if denominator == 0:
        return None
    return numerator / denominator * scale


# ---------------------------------------------------------------------------
# Рекламные данные
# ---------------------------------------------------------------------------


class _DataModel(BaseModel):
    """База моделей данных: алиасы Директа, неизменяемость, обрезка пробелов в строках."""

    model_config = ConfigDict(
        alias_generator=to_pascal,
        validate_by_name=True,
        validate_by_alias=True,
        frozen=True,
        str_strip_whitespace=True,
    )


class AdCounters(_DataModel):
    """Накопительные счётчики — общая часть суточной записи и агрегатов."""

    impressions: NonNegativeInt = Field(description="Показы")
    clicks: NonNegativeInt = Field(description="Клики")
    cost: Money = Field(description="Расход в валюте кабинета")
    conversions: NonNegativeInt = Field(description="Конверсии (лиды)")

    @field_validator("impressions", "clicks", "cost", "conversions", mode="before")
    @classmethod
    def _normalize_numeric_string(cls, value: Any) -> Any:
        """Строки из выгрузок: «1 234,50» → «1234.50», «--» (нет данных в API Директа) → 0."""
        if not isinstance(value, str):
            return value
        cleaned = value.strip()
        if cleaned == "--":
            return 0
        for thousands_separator in (" ", " ", " "):
            cleaned = cleaned.replace(thousands_separator, "")
        return cleaned.replace(",", ".")

    @model_validator(mode="after")
    def _check_clicks_within_impressions(self) -> Self:
        # Конверсии с кликами не сверяем: из-за окна атрибуции и нескольких целей
        # на визит их за сутки законно бывает больше, чем кликов.
        if self.clicks > self.impressions:
            raise ValueError(f"Кликов ({self.clicks}) больше, чем показов ({self.impressions})")
        return self


class DailyAdRecord(AdCounters):
    """Сырая суточная запись кампании — одна строка отчёта из кабинета.

    Лишние колонки выгрузки (например, Ctr или AvgCpc, которые считает сам Директ)
    игнорируются: производные метрики пересчитываются из сумм при агрегации.
    """

    date: dt.date = Field(description="Дата статистики")
    campaign_id: PositiveInt = Field(description="ID кампании в кабинете")
    campaign_name: str = Field(min_length=1, max_length=255, description="Название кампании")


class PerformanceMetrics(AdCounters):
    """Счётчики и производные KPI. Сама по себе — итог по набору кампаний.

    CTR и CR — в процентах (2.5 означает 2,5 %), CPC и CPA — в валюте кабинета.
    При нулевом знаменателе метрика равна None (см. ``safe_divide``).
    """

    @computed_field(alias="CTR")  # type: ignore[prop-decorator]
    @property
    def ctr(self) -> float | None:
        """CTR, % = клики / показы × 100."""
        return safe_divide(self.clicks, self.impressions, scale=100)

    @computed_field(alias="CPC")  # type: ignore[prop-decorator]
    @property
    def cpc(self) -> float | None:
        """Средняя цена клика = расход / клики."""
        return safe_divide(self.cost, self.clicks)

    @computed_field(alias="CPA")  # type: ignore[prop-decorator]
    @property
    def cpa(self) -> float | None:
        """Цена конверсии = расход / конверсии."""
        return safe_divide(self.cost, self.conversions)

    @computed_field(alias="CR")  # type: ignore[prop-decorator]
    @property
    def cr(self) -> float | None:
        """Конверсия из клика в лид, % = конверсии / клики × 100."""
        return safe_divide(self.conversions, self.clicks, scale=100)


class CampaignSummary(PerformanceMetrics):
    """Метрики одной кампании, агрегированные за период."""

    campaign_id: PositiveInt = Field(description="ID кампании в кабинете")
    campaign_name: str = Field(min_length=1, max_length=255, description="Название кампании")
    date_from: dt.date = Field(description="Первый день со статистикой")
    date_to: dt.date = Field(description="Последний день со статистикой")
    days_count: PositiveInt = Field(description="Число дней со статистикой")

    @model_validator(mode="after")
    def _check_period(self) -> Self:
        if self.date_from > self.date_to:
            raise ValueError(f"date_from ({self.date_from}) позже date_to ({self.date_to})")
        span = (self.date_to - self.date_from).days + 1
        if self.days_count > span:
            raise ValueError(f"days_count ({self.days_count}) больше длины периода ({span} дн.)")
        return self

    @classmethod
    def from_records(cls, records: Sequence[DailyAdRecord]) -> Self:
        """Сворачивает суточные записи одной кампании.

        KPI считаются от сумм (взвешенно), а не как среднее суточных значений:
        день с 10 показами не должен весить столько же, сколько день с 10 000.
        """
        if not records:
            raise ValueError("Нет записей для агрегации")
        campaign_ids = {record.campaign_id for record in records}
        if len(campaign_ids) > 1:
            raise ValueError(f"В одной группе записи разных кампаний: {sorted(campaign_ids)}")
        day_counts = Counter(record.date for record in records)
        duplicates = sorted(day for day, count in day_counts.items() if count > 1)
        if duplicates:
            raise ValueError(
                f"Кампания {records[0].campaign_id}: несколько записей за "
                f"{', '.join(map(str, duplicates))} — выгрузка загружена дважды?"
            )
        latest = max(records, key=lambda record: record.date)
        return cls(
            campaign_id=latest.campaign_id,
            campaign_name=latest.campaign_name,  # после переименования берём актуальное имя
            date_from=min(day_counts),
            date_to=latest.date,
            days_count=len(day_counts),
            **_sum_counters(records),
        )


# ---------------------------------------------------------------------------
# Агрегация
# ---------------------------------------------------------------------------


def _sum_counters(items: Iterable[AdCounters]) -> dict[str, Any]:
    """Суммы счётчиков; расход — через math.fsum, без накопления ошибки float."""
    rows = list(items)
    return {
        "impressions": sum(row.impressions for row in rows),
        "clicks": sum(row.clicks for row in rows),
        "cost": round(math.fsum(row.cost for row in rows), 2),
        "conversions": sum(row.conversions for row in rows),
    }


def aggregate_by_campaign(records: Iterable[DailyAdRecord]) -> list[CampaignSummary]:
    """Группирует суточные записи по campaign_id; сводки идут по убыванию расхода."""
    groups: defaultdict[int, list[DailyAdRecord]] = defaultdict(list)
    for record in records:
        groups[record.campaign_id].append(record)
    summaries = [CampaignSummary.from_records(group) for group in groups.values()]
    return sorted(summaries, key=lambda summary: (-summary.cost, summary.campaign_id))


def calculate_totals(items: Iterable[AdCounters]) -> PerformanceMetrics:
    """Итог по записям или сводкам: строка «ИТОГО», карточки KPI."""
    return PerformanceMetrics(**_sum_counters(items))


# ---------------------------------------------------------------------------
# Конфигурация отчёта (config/report_config.yaml)
# ---------------------------------------------------------------------------


class _ConfigModel(BaseModel):
    """База моделей конфига: неизвестный ключ в YAML — ошибка, а не молча пропущенная опечатка."""

    model_config = ConfigDict(extra="forbid", frozen=True, str_strip_whitespace=True)


def _ensure_unique(values: Iterable[str], what: str) -> None:
    duplicates = sorted(value for value, count in Counter(values).items() if count > 1)
    if duplicates:
        raise ValueError(f"Повторы в {what}: {', '.join(duplicates)}")


class ThemeConfig(_ConfigModel):
    """Оформление: цвета в формате RRGGBB (так их принимает python-pptx) и шрифты.

    Шрифт рисует PowerPoint получателя, поэтому по умолчанию — Calibri:
    он есть в любой установке Office и поддерживает кириллицу.
    """

    primary_color: str = Field(default="1A365D", pattern=_HEX_COLOR_PATTERN)
    accent_color: str = Field(default="00B4D8", pattern=_HEX_COLOR_PATTERN)
    bg_color: str = Field(default="F8F9FA", pattern=_HEX_COLOR_PATTERN)
    heading_font: str = Field(default="Calibri", min_length=1, max_length=64)
    body_font: str = Field(default="Calibri", min_length=1, max_length=64)

    @field_validator("primary_color", "accent_color", "bg_color", mode="before")
    @classmethod
    def _normalize_hex(cls, value: Any) -> Any:
        """«#1a365d» → «1A365D»."""
        return value.strip().removeprefix("#").upper() if isinstance(value, str) else value


class ReportMetadata(_ConfigModel):
    """Шапка отчёта: клиент, период, валюта, оформление."""

    client_name: str = Field(min_length=1, max_length=200)
    period_days: int = Field(default=30, ge=1, le=366, description="Длина периода, дней")
    date_to: dt.date | None = Field(
        default=None, description="Последний день периода включительно; None — вчера"
    )
    currency: str = Field(default="₽", min_length=1, max_length=5)
    theme: ThemeConfig = Field(default_factory=ThemeConfig)

    def resolve_period(self, today: dt.date | None = None) -> tuple[dt.date, dt.date]:
        """Границы периода включительно.

        Без явного date_to период заканчивается вчера — как пресеты LAST_N_DAYS
        в Reports API Директа: статистика за сегодня ещё неполная.
        """
        date_to = self.date_to or (today or dt.date.today()) - dt.timedelta(days=1)
        return date_to - dt.timedelta(days=self.period_days - 1), date_to


class _SlideBase(_ConfigModel):
    title: str = Field(min_length=1, max_length=120)
    enabled: bool = Field(default=True, description="false — слайд описан, но в отчёт не попадает")


class TitleSlide(_SlideBase):
    type: Literal["title_slide"]
    subtitle: str | None = None


class KpiMetric(_ConfigModel):
    id: KpiMetricId
    label: str = Field(min_length=1)
    format: MetricFormat


class KpiCardsSlide(_SlideBase):
    type: Literal["kpi_cards"]
    metrics: list[KpiMetric] = Field(min_length=1, max_length=8, description="Не больше 2 рядов по 4")

    @field_validator("metrics")
    @classmethod
    def _unique_metrics(cls, metrics: list[KpiMetric]) -> list[KpiMetric]:
        _ensure_unique((metric.id for metric in metrics), "metrics.id")
        return metrics


class TableColumn(_ConfigModel):
    field: SummaryField
    header: str = Field(min_length=1)


class SummaryTableSlide(_SlideBase):
    type: Literal["summary_table"]
    columns: list[TableColumn] = Field(min_length=1)
    highlight_metric: MetricField | None = None

    @model_validator(mode="after")
    def _check_columns(self) -> Self:
        fields = [column.field for column in self.columns]
        _ensure_unique(fields, "columns.field")
        if self.highlight_metric is not None and self.highlight_metric not in fields:
            raise ValueError(
                f"highlight_metric «{self.highlight_metric}» должен быть одной из колонок: "
                f"{', '.join(fields)}"
            )
        return self


class TrendChartSlide(_SlideBase):
    type: Literal["trend_chart"]
    chart_type: Literal["line", "bar"] = "line"
    x_axis: Literal["Date"] = "Date"
    series: list[MetricField] = Field(min_length=1)

    @field_validator("series")
    @classmethod
    def _unique_series(cls, series: list[str]) -> list[str]:
        _ensure_unique(series, "series")
        return series


class NotesSlide(_SlideBase):
    type: Literal["notes_slide"]
    placeholder_text: str = ""


Slide = Annotated[
    TitleSlide | KpiCardsSlide | SummaryTableSlide | TrendChartSlide | NotesSlide,
    Field(discriminator="type"),
]


class ReportConfig(_ConfigModel):
    """Конфигурация отчёта — структура config/report_config.yaml.

    Порядок слайдов и порядок метрик внутри слайда (metrics, columns, series)
    задаётся порядком элементов в YAML; повторы внутри списка запрещены.
    """

    report_metadata: ReportMetadata
    slides: list[Slide] = Field(min_length=1)

    @model_validator(mode="after")
    def _check_has_active_slides(self) -> Self:
        if not self.active_slides:
            raise ValueError("Все слайды выключены (enabled: false) — отчёт будет пустым")
        return self

    @property
    def active_slides(self) -> list[Slide]:
        """Слайды, которые попадут в презентацию, в порядке из конфига."""
        return [slide for slide in self.slides if slide.enabled]


# ---------------------------------------------------------------------------
# Выводы специалиста (слайд notes_slide)
# ---------------------------------------------------------------------------

NoteKind = Literal["heading", "paragraph", "bullet", "numbered"]

_NOTE_HEADING = re.compile(r"^\s{0,3}#{1,6}\s+(?P<text>.+?)\s*#*\s*$")
_NOTE_BULLET = re.compile(r"^\s*[-*•–—]\s+(?P<text>.+)$")
_NOTE_NUMBERED = re.compile(r"^\s*\d{1,3}[.)]\s+(?P<text>.+)$")
_NOTE_RULE = re.compile(r"^\s*([-*_=])(\s*\1){2,}\s*$")  # «---», «***» — разделители Markdown


class NoteBlock(_ConfigModel):
    kind: NoteKind
    text: str = Field(min_length=1, max_length=2000)


class ExpertNotes(_ConfigModel):
    """Выводы специалиста для слайда notes_slide.

    Текст разбирается как упрощённый Markdown: «# Заголовок», «- пункт» или «* пункт»,
    «1. пункт»; пустая строка разделяет абзацы, строки подряд склеиваются в один абзац
    или продолжают пункт списка. Фрагменты ``**так**`` рендерер выделяет полужирным.
    """

    blocks: list[NoteBlock] = Field(min_length=1, max_length=80)

    @classmethod
    def parse(cls, source: str | Sequence[str]) -> Self | None:
        """Строка — Markdown, последовательность строк — пункты списка. Пустой ввод — None."""
        if isinstance(source, str):
            blocks = _parse_note_markdown(source)
        else:
            blocks = [{"kind": "bullet", "text": " ".join(item.split())} for item in source if item.strip()]
        return cls.model_validate({"blocks": blocks}) if blocks else None


def _parse_note_markdown(text: str) -> list[dict[str, str]]:
    blocks: list[dict[str, str]] = []
    open_block: dict[str, str] | None = None  # блок, который продолжают следующие строки
    for line in text.splitlines():
        if not line.strip() or _NOTE_RULE.match(line):
            open_block = None
            continue
        if match := _NOTE_HEADING.match(line):
            blocks.append({"kind": "heading", "text": match["text"]})
            open_block = None
            continue
        if match := _NOTE_BULLET.match(line):
            open_block = {"kind": "bullet", "text": match["text"].strip()}
        elif match := _NOTE_NUMBERED.match(line):
            open_block = {"kind": "numbered", "text": match["text"].strip()}
        elif open_block is not None:
            open_block["text"] += " " + line.strip()
            continue
        else:
            open_block = {"kind": "paragraph", "text": line.strip()}
        blocks.append(open_block)
    return blocks
