"""Динамика метрик по дням на нативных диаграммах PowerPoint.

Каждая серия получает свою диаграмму с собственной осью: расход (десятки тысяч ₽)
и конверсии (единицы) на общей оси превратили бы конверсии в линию у нуля,
а вспомогательной оси в API python-pptx нет. Диаграммы стоят друг под другом
и делят ось дат, поэтому дни легко сопоставить.
"""

import math

from lxml import etree
from pptx.chart.chart import Chart
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_TICK_LABEL_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.slide import Slide
from pptx.util import Inches, Length, Pt

from src.rendering.context import RenderContext
from src.rendering.formatting import FIELD_SPECS, RATIO_FIELDS, field_title, format_field
from src.rendering.primitives import (
    CARD_PADDING,
    CONTENT_HEIGHT,
    CONTENT_TOP,
    CONTENT_WIDTH,
    GAP,
    MARGIN,
    add_card,
    add_slide_title,
    add_text,
)
from src.schemas import TrendChartSlide

_CHART_TYPES = {"line": XL_CHART_TYPE.LINE, "bar": XL_CHART_TYPE.COLUMN_CLUSTERED}
_HEADER_HEIGHT = Inches(0.35)
_DATE_FORMAT = "dd.mm"
_DATE_LABEL_WIDTH_IN = 0.7  # место под подпись «23.08» кеглем 10 с зазором


def render_trend_chart(slide: Slide, config: TrendChartSlide, ctx: RenderContext) -> None:
    theme = ctx.theme
    add_slide_title(slide, config.title, theme)

    count = len(config.series)
    columns = 1 if count <= 2 else 2
    rows = math.ceil(count / columns)
    cell_width = (CONTENT_WIDTH - GAP * (columns - 1)) // columns
    cell_height = (CONTENT_HEIGHT - GAP * (rows - 1)) // rows
    palette = (theme.primary, theme.accent)

    for index, field in enumerate(config.series):
        row, column = divmod(index, columns)
        left = MARGIN + column * (cell_width + GAP)
        top = CONTENT_TOP + row * (cell_height + GAP)
        add_card(slide, left, top, cell_width, cell_height, theme)
        _add_header(slide, field, ctx, left, top, cell_width)
        _add_chart(
            slide, config, field, ctx, palette[index % len(palette)],
            left, top + CARD_PADDING + _HEADER_HEIGHT, cell_width,
            cell_height - 2 * CARD_PADDING - _HEADER_HEIGHT,
        )  # fmt: skip


def _add_header(slide: Slide, field: str, ctx: RenderContext, left: int, top: int, width: int) -> None:
    """Название метрики слева, итог за период справа."""
    theme = ctx.theme
    inner_left, inner_width = left + CARD_PADDING, width - 2 * CARD_PADDING
    total = ctx.totals.model_dump(by_alias=True)[field]
    prefix = "В среднем" if field in RATIO_FIELDS else "За период"
    add_text(
        slide, inner_left, top + CARD_PADDING, inner_width // 2, _HEADER_HEIGHT, field_title(field, ctx.currency),
        font=theme.body_font, size=14, color=theme.text, bold=True, name=f"График {field}: название",
    )  # fmt: skip
    add_text(
        slide, inner_left + inner_width // 2, top + CARD_PADDING, inner_width - inner_width // 2,
        _HEADER_HEIGHT, f"{prefix}: {format_field(field, total)}",
        font=theme.body_font, size=12, color=theme.muted, align=PP_ALIGN.RIGHT,
        name=f"График {field}: итог",
    )  # fmt: skip


def _add_chart(
    slide: Slide,
    config: TrendChartSlide,
    field: str,
    ctx: RenderContext,
    color: RGBColor,
    left: int,
    top: int,
    width: int,
    height: int,
) -> None:
    spec = FIELD_SPECS[field]
    chart_data = CategoryChartData(number_format=spec.excel_format)
    chart_data.categories = [day for day, _ in ctx.daily]
    chart_data.categories.number_format = _DATE_FORMAT
    values = [metrics.model_dump(by_alias=True)[field] for _, metrics in ctx.daily]
    chart_data.add_series(
        field_title(field, ctx.currency),
        [None if value is None else round(value, spec.decimals or 2) for value in values],  # None — разрыв линии
    )

    chart_width = width - Inches(0.2)
    frame = slide.shapes.add_chart(
        _CHART_TYPES[config.chart_type], left + Inches(0.1), top, chart_width, height, chart_data
    )
    frame.name = f"График {field}"
    _style_chart(frame.chart, config, color, ctx)

    peak = max((value for value in values if value is not None), default=0)
    step = nice_step(peak)
    if spec.numeric and not spec.money and not spec.percent:
        step = max(step, 1)  # у счётчиков не бывает дробных делений
    value_axis = frame.chart.value_axis
    value_axis.major_unit = step
    value_axis.maximum_scale = step * (math.floor(peak / step) + 1)  # запас сверху, линия не упирается в рамку
    value_axis.tick_labels.number_format = axis_number_format(step, percent=spec.percent)
    value_axis.tick_labels.number_format_is_linked = False

    max_labels = max(2, int(Length(chart_width).inches / _DATE_LABEL_WIDTH_IN))
    _set_date_label_step(frame.chart, math.ceil(len(ctx.daily) / max_labels))


def nice_step(peak: float, target_ticks: int = 4) -> float:
    """Шаг делений вида 1, 2 или 5 × 10ⁿ — около ``target_ticks`` делений до пика."""
    if peak <= 0:
        return 1
    raw = peak / target_ticks
    magnitude = 10 ** math.floor(math.log10(raw))
    return next(multiplier * magnitude for multiplier in (1, 2, 5, 10) if raw <= multiplier * magnitude)


def axis_number_format(step: float, *, percent: bool) -> str:
    """Формат подписей оси, одинаковый при любой локали зрителя.

    Разделитель разрядов в «#,##0» PowerPoint берёт из региональных настроек
    («40,000» в en-US), поэтому разряды отделяются литеральными пробелами в условном
    формате. При целом шаге дробная часть не нужна — и десятичный разделитель тоже.
    """
    suffix = '"%"' if percent else ""
    if float(step).is_integer():
        return f"0{suffix}" if percent else "[>=1000000]0 000 000;[>=1000]0 000;0"
    decimals = 1 if float(step * 10).is_integer() else 2
    return f"0.{'0' * decimals}{suffix}"


def _set_date_label_step(chart: Chart, step_days: int) -> None:
    """Подписывает каждый N-й день, чтобы даты не наезжали друг на друга и не поворачивались.

    python-pptx не даёт API для шага оси дат, поэтому элементы c:majorUnit и
    c:majorTimeUnit добавляются в XML оси — сразу после c:baseTimeUnit, как требует схема.
    """
    axis = chart.category_axis._element  # noqa: SLF001 — публичного API для шага нет
    base_time_unit = axis.find(qn("c:baseTimeUnit"))
    if axis.tag != qn("c:dateAx") or base_time_unit is None or step_days <= 1:
        return
    major_time_unit = etree.Element(qn("c:majorTimeUnit"), val="days")
    major_unit = etree.Element(qn("c:majorUnit"), val=str(step_days))
    base_time_unit.addnext(major_time_unit)
    base_time_unit.addnext(major_unit)  # addnext вставляет вплотную: итог — majorUnit, majorTimeUnit


def _style_chart(chart: Chart, config: TrendChartSlide, color: RGBColor, ctx: RenderContext) -> None:
    theme = ctx.theme
    chart.has_legend = False
    chart.has_title = False  # название — в шапке карточки
    chart.font.name = theme.body_font
    chart.font.size = Pt(10)
    chart.font.color.rgb = theme.muted

    plot = chart.plots[0]
    series = plot.series[0]
    if config.chart_type == "line":
        series.smooth = False
        series.format.line.color.rgb = color
        series.format.line.width = Pt(2.25)
    else:
        plot.gap_width = 40
        plot.vary_by_categories = False
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = color

    value_axis = chart.value_axis
    value_axis.minimum_scale = 0  # метрики неотрицательны; ось от нуля не искажает динамику
    value_axis.has_major_gridlines = True
    value_axis.major_gridlines.format.line.color.rgb = theme.border
    value_axis.major_gridlines.format.line.width = Pt(0.75)
    value_axis.format.line.fill.background()

    category_axis = chart.category_axis
    category_axis.has_major_gridlines = False
    category_axis.format.line.color.rgb = theme.border
    category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    category_axis.tick_labels.number_format = _DATE_FORMAT
    category_axis.tick_labels.number_format_is_linked = False
