import datetime as dt
import json
from pathlib import Path

import openpyxl
import pytest

from src.data_loader import (
    SYNTHETIC_ID_BASE,
    campaign_id_from_name,
    format_summary_table,
    generate_mock_ad_data,
    load_ad_records,
    normalize_columns,
)
from src.schemas import CampaignSummary, aggregate_by_campaign, calculate_totals

END = dt.date(2026, 9, 21)
FIXTURES = Path(__file__).parent / "fixtures"
RU_MAPPING = {
    "Дата": "Date",
    "Направление": "CampaignName",
    "Показы": "Impressions",
    "Клики": "Clicks",
    "Затраты": "Cost",
    "Заявки": "Conversions",
}

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
    source = tmp_path / "report.xml"
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


# --- Excel и маппинг колонок ---------------------------------------------------


def write_xlsx(path: Path, rows: list[list], *, title: str = "Статистика", extra_sheets: int = 0) -> Path:
    workbook = openpyxl.Workbook()
    for index in range(extra_sheets):  # листы-обложки перед таблицей
        workbook.active.title = f"Обложка {index + 1}"
        workbook.active.append(["Отчёт для клиента"])
        workbook.create_sheet()
    sheet = workbook.active
    sheet.title = title
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    return path


RU_ROWS = [
    ["Статистика кампаний за август 2026"],
    [],
    ["Дата", "Направление", "Показы", "Клики", "Затраты", "Заявки", "CTR, %"],
    [dt.datetime(2026, 8, 24), "Поиск_Москва_Категории_Целевой", 5044, 127, 6111.23, 4, 2.52],
    [dt.datetime(2026, 8, 24), "РСЯ_РФ_Интересы_Аудитории", 15769, 321, 5017.65, 6, 2.04],
    [None, None, None, None, None, None, None],
    [dt.datetime(2026, 8, 25), "Поиск_Москва_Категории_Целевой", 4747, 115, 6381.98, 4, 2.42],
]


def test_xlsx_with_russian_headers_is_normalized_by_mapping(tmp_path):
    source = write_xlsx(tmp_path / "август.xlsx", RU_ROWS)

    records = load_ad_records(source, RU_MAPPING)

    assert len(records) == 3  # заголовок отчёта над таблицей и пустая строка пропущены
    first = records[0]
    assert first.date == dt.date(2026, 8, 24)  # datetime из Excel → дата
    assert first.campaign_name == "Поиск_Москва_Категории_Целевой"
    assert (first.impressions, first.clicks, first.cost, first.conversions) == (5044, 127, 6111.23, 4)
    assert first.campaign_id == campaign_id_from_name("Поиск_Москва_Категории_Целевой")
    assert [s.days_count for s in aggregate_by_campaign(records)] == [2, 1]


def test_xlsx_in_direct_notation_needs_no_mapping(tmp_path):
    source = write_xlsx(
        tmp_path / "direct.xlsx",
        [
            ["Date", "CampaignId", "CampaignName", "Impressions", "Clicks", "Cost", "Conversions"],
            ["2026-08-24", 98743239, "Поиск_Бренд", 4126, 158, "5 964,74", "--"],
        ],
        extra_sheets=1,
    )

    [record] = load_ad_records(source)

    assert record.campaign_id == 98743239
    assert (record.date, record.cost, record.conversions) == (dt.date(2026, 8, 24), 5964.74, 0)


def test_legacy_xls_is_read_with_mapping():
    """Фикстура — книга Excel 97–2003: заголовок отчёта, пустая строка, таблица с русскими колонками."""
    records = load_ad_records(FIXTURES / "direct_export_ru.xls", RU_MAPPING)

    assert [record.date for record in records] == [dt.date(2026, 8, 24), dt.date(2026, 8, 24), dt.date(2026, 8, 25)]
    assert records[1].campaign_name == "РСЯ_РФ_Интересы_Аудитории"
    assert (records[1].impressions, records[1].cost, records[1].conversions) == (15769, 5017.65, 6)


def test_excel_format_is_detected_by_content_not_extension(tmp_path):
    renamed = write_xlsx(tmp_path / "export.xls", RU_ROWS)  # .xlsx, сохранённый с расширением .xls

    assert len(load_ad_records(renamed, RU_MAPPING)) == 3


def test_not_an_excel_file_is_reported(tmp_path):
    source = tmp_path / "export.xlsx"
    source.write_text("<html><table><tr><td>Дата</td></tr></table></html>", encoding="utf-8")

    with pytest.raises(ValueError, match="не похож на книгу Excel"):
        load_ad_records(source)


def test_corrupted_xlsx_is_reported(tmp_path):
    source = tmp_path / "broken.xlsx"
    source.write_bytes(b"PK\x03\x04" + b"\x00" * 64)

    with pytest.raises(ValueError, match="не удалось прочитать книгу Excel"):
        load_ad_records(source)


def test_missing_columns_point_to_column_mapping(tmp_path):
    source = write_xlsx(tmp_path / "export.xlsx", RU_ROWS)

    with pytest.raises(ValueError, match=r"нет обязательных колонок Date, CampaignName.*column_mapping"):
        load_ad_records(source)  # русские колонки без маппинга


def test_duplicate_excel_headers_are_rejected(tmp_path):
    rows = [["Дата", "Направление", "Показы", "Клики", "Затраты", "Затраты", "Заявки"]]
    source = write_xlsx(tmp_path / "export.xlsx", rows)

    with pytest.raises(ValueError, match="повторяются колонки Затраты"):
        load_ad_records(source, RU_MAPPING)


def test_empty_workbook_gives_no_records(tmp_path):
    assert load_ad_records(write_xlsx(tmp_path / "empty.xlsx", []), RU_MAPPING) == []


def test_csv_with_mapping_and_russian_dates(tmp_path):
    source = tmp_path / "export.csv"
    source.write_text(
        "Дата;Направление;Показы;Клики;Затраты;Заявки\n24.08.2026;Поиск_Бренд;4 126;158;5964,74;3\n",
        encoding="utf-8-sig",
    )

    [record] = load_ad_records(source, RU_MAPPING)

    assert record.date == dt.date(2026, 8, 24)
    assert (record.impressions, record.cost, record.conversions) == (4126, 5964.74, 3)


def test_json_with_mapping(tmp_path):
    source = tmp_path / "export.json"
    rows = [{"Дата": "2026-08-24", "Направление": "РСЯ_Тест", "Показы": 100, "Клики": 5, "Затраты": 50.5, "Заявки": 1}]
    source.write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")

    [record] = load_ad_records(source, RU_MAPPING)

    assert (record.campaign_name, record.cost) == ("РСЯ_Тест", 50.5)


def test_normalize_columns_ignores_case_and_extra_spaces():
    row = {"  затраты ": "10", "НАПРАВЛЕНИЕ": "Поиск", "Clicks": 3, "CTR, %": 1.5}

    assert normalize_columns(row, {"Затраты": "Cost", "Направление": "CampaignName"}) == {
        "Cost": "10",
        "CampaignName": "Поиск",
        "Clicks": 3,  # уже в нотации Директа — без изменений
        "CTR, %": 1.5,  # лишние колонки остаются, DailyAdRecord их игнорирует
    }
    assert normalize_columns(row, None) == row
    assert normalize_columns(row, {}) == row


def test_normalize_columns_rejects_two_sources_for_one_column():
    with pytest.raises(ValueError, match="«Cost» и «Затраты» обе соответствуют Cost"):
        normalize_columns({"Cost": 1, "Затраты": 2}, {"Затраты": "Cost"})


def test_mapping_conflict_is_reported_with_row_number(tmp_path):
    source = tmp_path / "export.csv"
    with_both = DIRECT_CSV.replace("Conversions", "Conversions,Затраты").replace("\n2026", ",1\n2026")
    source.write_text(with_both, encoding="utf-8")

    with pytest.raises(ValueError, match="запись 1: колонки «Cost» и «Затраты»"):
        load_ad_records(source, {"Затраты": "Cost"})
