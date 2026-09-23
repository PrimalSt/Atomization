"""Форматирование значений для слайдов: русские разделители, валюта, проценты."""

from dataclasses import dataclass
from typing import Any, get_args

from src.schemas import KpiMetricId, MetricFormat, PerformanceMetrics, RatioField

NBSP = " "  # неразрывный пробел: PowerPoint не разорвёт число переносом строки
MISSING = "—"  # метрика не определена: нулевой знаменатель
RATIO_FIELDS: frozenset[str] = frozenset(get_args(RatioField))  # CTR, CPC, CPA, CR — не суммируются


def format_number(value: float, decimals: int = 0) -> str:
    """12345.6 → «12 345,60»: пробел между разрядами, запятая перед дробной частью."""
    return f"{value:,.{decimals}f}".replace(",", NBSP).replace(".", ",")


def format_value(value: float | None, fmt: MetricFormat, currency: str, decimals: int = 0) -> str:
    """Значение KPI в формате из конфига. CTR и CR уже в процентах: 2.5 → «2,50%»."""
    if value is None:
        return MISSING
    match fmt:
        case "currency":
            return f"{format_number(value, decimals)}{NBSP}{currency}"
        case "percent":
            return f"{format_number(value, decimals)}%"
        case "integer":
            return format_number(value, 0)
        case "decimal":
            return format_number(value, decimals)
    raise ValueError(f"Неизвестный формат значения: {fmt}")


@dataclass(frozen=True, slots=True)
class FieldSpec:
    """Как показывать поле сводки (ключ — алиас из SummaryField)."""

    label: str
    decimals: int = 0
    percent: bool = False
    money: bool = False
    numeric: bool = True

    @property
    def excel_format(self) -> str:
        """Формат значений во встроенной таблице диаграммы (разделители — по локали зрителя)."""
        number = "#,##0" + ("." + "0" * self.decimals if self.decimals else "")
        return f'{number}"%"' if self.percent else number


FIELD_SPECS: dict[str, FieldSpec] = {
    "CampaignId": FieldSpec("ID кампании", numeric=False),
    "CampaignName": FieldSpec("Кампания", numeric=False),
    "DaysCount": FieldSpec("Дней"),
    "Impressions": FieldSpec("Показы"),
    "Clicks": FieldSpec("Клики"),
    "Cost": FieldSpec("Расход", money=True),
    "Conversions": FieldSpec("Конверсии"),
    "CTR": FieldSpec("CTR", decimals=2, percent=True),
    "CPC": FieldSpec("CPC", decimals=2, money=True),
    "CPA": FieldSpec("CPA", money=True),
    "CR": FieldSpec("CR", decimals=2, percent=True),
}


def format_field(field: str, value: Any) -> str:
    """Значение поля сводки без символа валюты (он в заголовке колонки или графика)."""
    if value is None:
        return MISSING
    if isinstance(value, str):
        return value
    spec = FIELD_SPECS[field]
    if not spec.numeric:
        return str(value)
    text = format_number(value, spec.decimals)
    return f"{text}%" if spec.percent else text


def field_title(field: str, currency: str) -> str:
    """«Расход, ₽», «CTR, %», «Клики»."""
    spec = FIELD_SPECS[field]
    if spec.money:
        return f"{spec.label}, {currency}"
    if spec.percent:
        return f"{spec.label}, %"
    return spec.label


@dataclass(frozen=True, slots=True)
class KpiSource:
    """Откуда брать значение KPI и что написать под ним.

    ``formula`` задана у производных метрик; для накопительных под значением
    показывается среднее за день.
    """

    attribute: str  # поле или вычисляемое поле PerformanceMetrics
    decimals: int = 0
    formula: str | None = None


KPI_SOURCES: dict[KpiMetricId, KpiSource] = {
    "total_spend": KpiSource("cost"),
    "impressions": KpiSource("impressions"),
    "clicks": KpiSource("clicks"),
    "conversions": KpiSource("conversions"),
    "avg_ctr": KpiSource("ctr", decimals=2, formula="клики ÷ показы"),
    "avg_cpc": KpiSource("cpc", decimals=2, formula="расход ÷ клики"),
    "avg_cpa": KpiSource("cpa", formula="расход ÷ конверсии"),
    "avg_cr": KpiSource("cr", decimals=2, formula="конверсии ÷ клики"),
}


def kpi_value(metrics: PerformanceMetrics, metric_id: KpiMetricId) -> float | None:
    return getattr(metrics, KPI_SOURCES[metric_id].attribute)
