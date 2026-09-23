import datetime as dt
import json

import pytest

from src.data_loader import (
    SYNTHETIC_ID_BASE,
    campaign_id_from_name,
    format_summary_table,
    generate_mock_ad_data,
    load_ad_records,
)
from src.schemas import CampaignSummary, aggregate_by_campaign, calculate_totals

END = dt.date(2026, 9, 21)

DIRECT_CSV = """Date,CampaignName,Impressions,Clicks,Cost,Conversions
2026-08-24,Поиск_Москва_Категории_Целевой,5044,127,6111.23,4
2026-08-24,РСЯ_РФ_Интересы_Аудитории,15769,321,5017.65,6
2026-08-25,Поиск_Москва_Категории_Целевой,4747,115,6381.98,4
"""


def test_csv_without_campaign_id_gets_stable_synthetic_ids(tmp_path):
    source = tmp_path / "direct.csv"
    source.write_text(DIRECT_CSV, encoding="utf-8")

    records = load_ad_records(source)

    assert len(records) == 3
    assert records[0].campaign_id == records[2].campaign_id == campaign_id_from_name("Поиск_Москва_Категории_Целевой")
    assert records[0].campaign_id != records[1].campaign_id
    assert all(record.campaign_id >= SYNTHETIC_ID_BASE for record in records)
    assert load_ad_records(source) == records  # тот же файл — те же ID
    assert [s.days_count for s in aggregate_by_campaign(records)] == [2, 1]


def test_excel_csv_with_semicolons_and_comma_decimals(tmp_path):
    source = tmp_path / "excel.csv"
    source.write_text(
        "Date;CampaignId;CampaignName;Impressions;Clicks;Cost;Conversions\n"
        "2026-08-24;98743239;Поиск_Бренд;4 126;158;5964,74;--\n",
        encoding="utf-8-sig",
    )

    [record] = load_ad_records(source)

    assert record.campaign_id == 98743239  # реальный ID не подменяется
    assert (record.impressions, record.cost, record.conversions) == (4126, 5964.74, 0)


def test_json_export(tmp_path):
    source = tmp_path / "direct.json"
    rows = [{"Date": "2026-08-24", "CampaignName": "РСЯ_Тест", "Impressions": 100, "Clicks": 5, "Cost": 50.5, "Conversions": 1}]
    source.write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")

    [record] = load_ad_records(source)

    assert record.campaign_name == "РСЯ_Тест"
    assert record.campaign_id == campaign_id_from_name("РСЯ_Тест")


def test_invalid_row_is_reported_with_its_number(tmp_path):
    source = tmp_path / "broken.csv"
    source.write_text(DIRECT_CSV + "2026-08-26,Поиск_Тест,10,50,100.0,1\n", encoding="utf-8")

    with pytest.raises(ValueError, match="запись 4"):
        load_ad_records(source)


def test_unsupported_format(tmp_path):
    source = tmp_path / "report.xlsx"
    source.write_bytes(b"")

    with pytest.raises(ValueError, match="неподдерживаемый формат"):
        load_ad_records(source)


def test_generates_full_grid_of_days_and_campaigns():
    records = generate_mock_ad_data(days=30, campaigns_count=3, end_date=END, seed=1)

    assert len(records) == 90
    assert len({record.campaign_id for record in records}) == 3
    assert min(record.date for record in records) == END - dt.timedelta(days=29)
    assert max(record.date for record in records) == END
    assert records == sorted(records, key=lambda record: record.date)


def test_default_period_ends_yesterday():
    records = generate_mock_ad_data(days=1, seed=1)

    assert {record.date for record in records} == {dt.date.today() - dt.timedelta(days=1)}


def test_same_seed_gives_same_data():
    assert generate_mock_ad_data(seed=7, end_date=END) == generate_mock_ad_data(seed=7, end_date=END)
    assert generate_mock_ad_data(seed=7, end_date=END) != generate_mock_ad_data(seed=8, end_date=END)


@pytest.mark.parametrize("seed", range(10))
def test_campaign_metrics_are_realistic(seed):
    summaries = aggregate_by_campaign(generate_mock_ad_data(campaigns_count=8, end_date=END, seed=seed))

    for summary in summaries:
        assert 1.5 <= summary.ctr <= 3.5, summary.campaign_name
        assert 2.0 <= summary.cr <= 5.0, summary.campaign_name


def test_names_stay_unique_beyond_template_pool():
    records = generate_mock_ad_data(days=1, campaigns_count=10, end_date=END, seed=1)

    assert len({record.campaign_name for record in records}) == 10


@pytest.mark.parametrize("kwargs", [{"days": 0}, {"campaigns_count": 0}])
def test_rejects_empty_period_or_no_campaigns(kwargs):
    with pytest.raises(ValueError):
        generate_mock_ad_data(**kwargs)


def test_summary_table_shows_totals_and_undefined_metrics():
    idle = CampaignSummary(
        campaign_id=1,
        campaign_name="Пауза",
        date_from=END,
        date_to=END,
        days_count=1,
        impressions=0,
        clicks=0,
        cost=0.0,
        conversions=0,
    )

    table = format_summary_table([idle], calculate_totals([idle]), currency="₽")

    lines = table.splitlines()
    assert "CPA, ₽" in lines[0]
    assert lines[2].startswith("Пауза") and lines[2].count("—") == 4
    assert lines[-1].startswith("ИТОГО")
    assert len({len(line) for line in lines}) == 1  # колонки выровнены
