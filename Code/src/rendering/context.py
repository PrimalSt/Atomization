"""Данные, общие для всех рендереров слайдов."""

import datetime as dt
from collections import defaultdict
from collections.abc import Sequence
from dataclasses import dataclass

from src.rendering.theme import Theme
from src.schemas import (
    CampaignSummary,
    DailyAdRecord,
    ExpertNotes,
    PerformanceMetrics,
    ReportConfig,
    calculate_totals,
)


@dataclass(frozen=True, slots=True)
class RenderContext:
    config: ReportConfig
    theme: Theme
    summaries: tuple[CampaignSummary, ...]
    totals: PerformanceMetrics  # итог по кампаниям: карточки KPI и строка «ИТОГО»
    daily: tuple[tuple[dt.date, PerformanceMetrics], ...]  # итоги по дням, даты без пропусков
    period: tuple[dt.date, dt.date]
    generated_at: dt.date
    notes: ExpertNotes | None = None  # выводы специалиста; None — на слайде останется заготовка
    notes_layout: tuple[int, int] | None = None  # (кегль, колонок), общие для всех слайдов разбитых выводов

    @classmethod
    def from_data(
        cls,
        config: ReportConfig,
        summaries: Sequence[CampaignSummary],
        daily_records: Sequence[DailyAdRecord],
        generated_at: dt.date,
        notes: ExpertNotes | None = None,
    ) -> "RenderContext":
        """Готовит данные для слайдов.

        Период берётся из самих записей — отчёт показывает то, что реально загружено.
        Если записей нет, используется период из конфига. Дни без статистики
        заполняются нулями, чтобы на графике не было «слипшихся» дат.
        """
        records_by_day: defaultdict[dt.date, list[DailyAdRecord]] = defaultdict(list)
        for record in daily_records:
            records_by_day[record.date].append(record)
        if records_by_day:
            date_from, date_to = min(records_by_day), max(records_by_day)
        else:
            date_from, date_to = config.report_metadata.resolve_period(today=generated_at)

        days = [date_from + dt.timedelta(days=offset) for offset in range((date_to - date_from).days + 1)]
        return cls(
            config=config,
            theme=Theme.from_config(config.report_metadata.theme),
            summaries=tuple(summaries),
            totals=calculate_totals(summaries),
            daily=tuple((day, calculate_totals(records_by_day.get(day, ()))) for day in days),
            period=(date_from, date_to),
            generated_at=generated_at,
            notes=notes,
        )

    @property
    def client_name(self) -> str:
        return self.config.report_metadata.client_name

    @property
    def currency(self) -> str:
        return self.config.report_metadata.currency

    @property
    def period_label(self) -> str:
        date_from, date_to = self.period
        return f"{date_from:%d.%m.%Y} – {date_to:%d.%m.%Y}"
