import datetime as dt
import json
from typing import get_args

import pytest
from pydantic import ValidationError

from src.config_loader import DEFAULT_CONFIG_PATH, load_report_config
from src.schemas import (
    CampaignSummary,
    DailyAdRecord,
    ExpertNotes,
    KpiCardsSlide,
    ReportConfig,
    SummaryField,
    aggregate_by_campaign,
    calculate_totals,
)

DAY = dt.date(2026, 9, 1)


def make_record(day: dt.date = DAY, **overrides) -> DailyAdRecord:
    values = {
        "date": day,
        "campaign_id": 101,
        "campaign_name": "Поиск_Тест",
        "impressions": 1000,
        "clicks": 50,
        "cost": 2500.0,
        "conversions": 5,
    }
    return DailyAdRecord(**(values | overrides))


def make_summary(**overrides) -> CampaignSummary:
    values = {
        "campaign_id": 101,
        "campaign_name": "Поиск_Тест",
        "date_from": DAY,
        "date_to": DAY,
        "days_count": 1,
        "impressions": 1000,
        "clicks": 50,
        "cost": 2500.0,
        "conversions": 5,
    }
    return CampaignSummary(**(values | overrides))


def test_daily_record_accepts_direct_export_row():
    record = DailyAdRecord.model_validate(
        {
            "Date": "2026-09-01",
            "CampaignId": "98743239",
            "CampaignName": "  Поиск_РФ_Бренд  ",
            "Impressions": "12 345",
            "Clicks": "321",
            "Cost": "6 111,23",
            "Conversions": "--",
            "Ctr": "2.60",  # лишние колонки выгрузки игнорируются
        }
    )

    assert record.date == DAY
    assert record.campaign_name == "Поиск_РФ_Бренд"
    assert record.impressions == 12345
    assert record.cost == pytest.approx(6111.23)
    assert record.conversions == 0


@pytest.mark.parametrize(
    "overrides",
    [
        {"clicks": 1001},
        {"clicks": -1},
        {"cost": -0.01},
        {"cost": float("nan")},
        {"cost": float("inf")},
        {"campaign_name": "   "},
        {"campaign_id": 0},
        {"impressions": 10.5},
    ],
)
def test_daily_record_rejects_invalid_values(overrides):
    with pytest.raises(ValidationError):
        make_record(**overrides)


def test_metrics_formulas():
    summary = make_summary()

    assert summary.ctr == pytest.approx(5.0)
    assert summary.cpc == pytest.approx(50.0)
    assert summary.cpa == pytest.approx(500.0)
    assert summary.cr == pytest.approx(10.0)


@pytest.mark.parametrize(
    ("counters", "undefined"),
    [
        ({"impressions": 0, "clicks": 0, "cost": 0.0, "conversions": 0}, {"ctr", "cpc", "cpa", "cr"}),
        ({"clicks": 0, "cost": 0.0, "conversions": 0}, {"cpc", "cpa", "cr"}),
        ({"conversions": 0}, {"cpa"}),
    ],
)
def test_zero_denominators_give_none_instead_of_error(counters, undefined):
    summary = make_summary(**counters)

    dumped = summary.model_dump()
    assert {name for name in ("ctr", "cpc", "cpa", "cr") if dumped[name] is None} == undefined
    json.loads(summary.model_dump_json())  # без inf/NaN JSON остаётся валидным


def test_summary_round_trips_through_its_own_dump():
    summary = make_summary()

    assert CampaignSummary.model_validate(summary.model_dump(by_alias=True)) == summary


def test_summary_fields_used_by_config_exist_in_dump():
    dumped = make_summary().model_dump(by_alias=True)

    assert set(get_args(SummaryField)) <= dumped.keys()


def test_aggregation_uses_weighted_metrics():
    records = [
        make_record(DAY, impressions=100, clicks=10, cost=100.0, conversions=1),
        make_record(DAY + dt.timedelta(days=2), impressions=10_000, clicks=100, cost=1000.0, conversions=0),
        make_record(DAY, campaign_id=202, campaign_name="РСЯ_Тест", impressions=500, clicks=5, cost=50.0, conversions=0),
    ]

    summaries = aggregate_by_campaign(records)
    search = summaries[0]

    assert [s.campaign_id for s in summaries] == [101, 202]  # по убыванию расхода
    assert search.ctr == pytest.approx(110 / 10_100 * 100)  # не среднее (10 % + 1 %) / 2
    assert (search.date_from, search.date_to, search.days_count) == (DAY, DAY + dt.timedelta(days=2), 2)
    assert summaries[1].cpa is None
    totals = calculate_totals(summaries)
    assert (totals.impressions, totals.cost, totals.conversions) == (10_600, 1150.0, 1)


def test_aggregation_takes_latest_campaign_name():
    records = [make_record(DAY, campaign_name="Старое"), make_record(DAY + dt.timedelta(days=1), campaign_name="Новое")]

    assert aggregate_by_campaign(records)[0].campaign_name == "Новое"


def test_aggregation_rejects_duplicate_days():
    with pytest.raises(ValueError, match="несколько записей"):
        aggregate_by_campaign([make_record(), make_record()])


def test_summary_rejects_inconsistent_period():
    with pytest.raises(ValidationError):
        make_summary(date_from=DAY, date_to=DAY - dt.timedelta(days=1))
    with pytest.raises(ValidationError):
        make_summary(days_count=2)


def test_project_config_is_valid():
    config = load_report_config(DEFAULT_CONFIG_PATH)

    assert config.report_metadata.client_name == "ООО Доставка Плюс"
    assert [slide.type for slide in config.active_slides] == [
        "title_slide", "kpi_cards", "summary_table", "trend_chart", "notes_slide",
    ]  # fmt: skip
    kpi = config.active_slides[1]
    assert isinstance(kpi, KpiCardsSlide)
    assert [metric.id for metric in kpi.metrics] == ["total_spend", "conversions", "avg_cpa", "avg_ctr"]
    assert config.report_metadata.resolve_period(today=dt.date(2026, 9, 22)) == (
        dt.date(2026, 8, 23),
        dt.date(2026, 9, 21),
    )


def base_config() -> dict:
    return {
        "report_metadata": {"client_name": "Клиент", "theme": {"primary_color": "#1a365d"}},
        "slides": [
            {"type": "kpi_cards", "title": "KPI", "metrics": [{"id": "avg_cpa", "label": "CPA", "format": "currency"}]},
            {"type": "notes_slide", "title": "Выводы", "enabled": False},
        ],
    }


def test_config_normalizes_colors_and_filters_disabled_slides():
    config = ReportConfig.model_validate(base_config())

    assert config.report_metadata.theme.primary_color == "1A365D"
    assert [slide.type for slide in config.active_slides] == ["kpi_cards"]


@pytest.mark.parametrize(
    "mutate",
    [
        lambda c: c["slides"][0]["metrics"][0].update(id="avg_roi"),  # неизвестная метрика
        lambda c: c["slides"][0]["metrics"].append(dict(c["slides"][0]["metrics"][0])),  # повтор метрики
        lambda c: c["slides"].append({"type": "pie_chart", "title": "?"}),  # неизвестный тип слайда
        lambda c: c["report_metadata"].update(clinet_name="опечатка в ключе"),
        lambda c: c["report_metadata"].update(period_days=0),
        lambda c: c["report_metadata"]["theme"].update(bg_color="white"),
        lambda c: c["slides"][0].update(enabled=False),  # не осталось активных слайдов
        lambda c: c["slides"].append(
            {"type": "summary_table", "title": "Т", "columns": [{"field": "Cost", "header": "Расход"}], "highlight_metric": "CPA"}
        ),  # подсвечивается метрика, которой нет в таблице
    ],
)
def test_config_rejects_invalid_values(mutate):
    raw = base_config()
    mutate(raw)

    with pytest.raises(ValidationError):
        ReportConfig.model_validate(raw)


def blocks(notes: ExpertNotes) -> list[tuple[str, str]]:
    return [(block.kind, block.text) for block in notes.blocks]


def test_expert_notes_parse_markdown():
    notes = ExpertNotes.parse(
        "# Итоги\n\nРасход вырос,\nCPA стабилен.\n\n---\n- Отключили площадки\n  с высоким CPA\n* Добавили минус-слова\n"
        "1. Тест объявлений\n2) Снизить ставки\n## План ##\n— Подключить фид"
    )

    assert blocks(notes) == [
        ("heading", "Итоги"),
        ("paragraph", "Расход вырос, CPA стабилен."),
        ("bullet", "Отключили площадки с высоким CPA"),
        ("bullet", "Добавили минус-слова"),
        ("numbered", "Тест объявлений"),
        ("numbered", "Снизить ставки"),
        ("heading", "План"),
        ("bullet", "Подключить фид"),
    ]


def test_expert_notes_from_items_and_blank_input():
    assert blocks(ExpertNotes.parse(["  Первый\n пункт ", "", "Второй"])) == [
        ("bullet", "Первый пункт"),
        ("bullet", "Второй"),
    ]
    assert ExpertNotes.parse("  \n\n ") is None
    assert ExpertNotes.parse([]) is None


def test_expert_notes_limit_block_length():
    with pytest.raises(ValidationError):
        ExpertNotes.parse("x" * 2001)