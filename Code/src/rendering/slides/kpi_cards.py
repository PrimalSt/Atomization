"""Карточки KPI: крупное значение, подпись и пояснение под ним.

Пояснение — динамика к прошлому месяцу, если он есть в локальной истории;
иначе формула (для производных метрик) или среднее за день (для накопительных).
"""

import math

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.slide import Slide
from pptx.util import Inches, Length

from src.rendering.context import RenderContext
from src.rendering.formatting import KPI_SOURCES, format_value, kpi_value, month_over_month
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
from src.schemas import KpiCardsSlide, KpiMetric

_MAX_COLUMNS = 4
_VALUE_MAX_PT = 54
_VALUE_MIN_PT = 20
_CHAR_WIDTH_EM = 0.58  # средняя ширина символа жирных цифр Calibri с запасом


def render_kpi_cards(slide: Slide, config: KpiCardsSlide, ctx: RenderContext) -> None:
    add_slide_title(slide, config.title, ctx.theme)

    count = len(config.metrics)
    columns = count if count <= _MAX_COLUMNS else math.ceil(count / 2)
    rows = math.ceil(count / columns)
    card_width = (CONTENT_WIDTH - GAP * (columns - 1)) // columns
    card_height = Inches(3.0) if rows == 1 else Inches(2.35)
    grid_height = rows * card_height + (rows - 1) * GAP
    grid_top = CONTENT_TOP + (CONTENT_HEIGHT - grid_height) // 2

    values = [kpi_value(ctx.totals, metric.id) for metric in config.metrics]
    texts = [
        format_value(value, metric.format, ctx.currency, KPI_SOURCES[metric.id].decimals)
        for metric, value in zip(config.metrics, values)
    ]
    # Один кегль на все карточки: разный размер цифр читается как разная важность.
    value_size = min(_fit_font_size(text, card_width - 2 * CARD_PADDING) for text in texts)

    for index, (metric, value, text) in enumerate(zip(config.metrics, values, texts)):
        row, column = divmod(index, columns)
        cards_in_row = min(columns, count - row * columns)
        row_shift = (columns - cards_in_row) * (card_width + GAP) // 2  # неполный ряд — по центру
        left = MARGIN + row_shift + column * (card_width + GAP)
        top = grid_top + row * (card_height + GAP)
        _render_card(slide, metric, value, text, value_size, ctx, (left, top, card_width, card_height))


def _render_card(
    slide: Slide,
    metric: KpiMetric,
    value: float | None,
    value_text: str,
    value_size: float,
    ctx: RenderContext,
    box: tuple[int, int, int, int],
) -> None:
    theme = ctx.theme
    left, top, width, height = box
    add_card(slide, left, top, width, height, theme)
    inner_left, inner_width = left + CARD_PADDING, width - 2 * CARD_PADDING
    add_text(
        slide, inner_left, top + CARD_PADDING, inner_width, Inches(0.7), metric.label,
        font=theme.body_font, size=16, color=theme.muted, bold=True, name=f"KPI {metric.id}: подпись",
    )  # fmt: skip
    value_height = Inches(1.0)
    value_top = top + (height - value_height) // 2 + Inches(0.1)
    add_text(
        slide, inner_left, value_top, inner_width, value_height, value_text,
        font=theme.heading_font, size=value_size, color=theme.primary,
        bold=True, anchor=MSO_ANCHOR.MIDDLE, name=f"KPI {metric.id}: значение",
    )  # fmt: skip
    caption, caption_color = _caption(metric, value, ctx)
    add_text(
        slide, inner_left, top + height - CARD_PADDING - Inches(0.35), inner_width, Inches(0.35),
        caption,
        font=theme.body_font, size=13, color=caption_color, anchor=MSO_ANCHOR.BOTTOM,
        name=f"KPI {metric.id}: пояснение",
    )  # fmt: skip


def _fit_font_size(text: str, width: int) -> float:
    """Самый крупный кегль, при котором значение помещается в одну строку карточки."""
    fitting = Length(width).pt / (_CHAR_WIDTH_EM * max(len(text), 1))
    return max(_VALUE_MIN_PT, min(_VALUE_MAX_PT, math.floor(fitting)))


def _caption(metric: KpiMetric, value: float | None, ctx: RenderContext) -> tuple[str, RGBColor]:
    """Текст и цвет пояснения: динамика MoM, если её можно посчитать, иначе справка о метрике."""
    if ctx.comparison is not None:
        previous = kpi_value(ctx.comparison.metrics, metric.id)
        change = month_over_month(value, previous, metric.format, ctx.comparison.month)
        if change is not None:
            return change.text, _change_color(metric, change.direction, ctx)
    return _reference_caption(metric, value, ctx), ctx.theme.muted


def _change_color(metric: KpiMetric, direction: int, ctx: RenderContext) -> RGBColor:
    """Зелёный — изменение к лучшему, красный — к худшему; расход и «без изменений» нейтральны."""
    better = KPI_SOURCES[metric.id].better
    if better is None or direction == 0:
        return ctx.theme.muted
    improved = (direction > 0) == (better == "higher")
    return ctx.theme.good if improved else ctx.theme.bad


def _reference_caption(metric: KpiMetric, value: float | None, ctx: RenderContext) -> str:
    """Формула для производных метрик, среднее за день — для накопительных."""
    source = KPI_SOURCES[metric.id]
    if source.formula:
        return source.formula
    if value is None or not ctx.daily:
        return ""
    per_day = format_value(value / len(ctx.daily), metric.format, ctx.currency, source.decimals)
    return f"≈ {per_day} в день"
