"""Логика окна без tkinter: проверка ввода, рабочий поток, конфиги первого запуска .exe."""

import logging
import queue
import subprocess
import sys
import threading
from pathlib import Path

import pytest

from src.config_loader import DEFAULT_CONFIG_PATH, PROJECT_ROOT
from src.gui.tasks import (
    FailureEvent,
    GenerationRequest,
    LogEvent,
    ProgressEvent,
    QueueLogHandler,
    SuccessEvent,
    describe_config,
    list_config_files,
    read_notes_file,
    run_generation,
    validate_request,
)
from src.paths import ensure_user_configs


def request(tmp_path: Path, **overrides) -> GenerationRequest:
    values = {
        "config_path": DEFAULT_CONFIG_PATH,
        "input_path": None,
        "mock": True,
        "notes": "",
        "output_dir": tmp_path / "output",
        "history_db": tmp_path / "history.db",
    }
    return GenerationRequest(**(values | overrides))


def drain(events: queue.Queue) -> list:
    items = []
    while not events.empty():
        items.append(events.get_nowait())
    return items


def test_gui_logic_does_not_import_tkinter():
    """Логика окна импортируется без tkinter: её тесты идут и там, где нет графической среды."""
    code = "import sys, src.gui.tasks; print(sorted({'tkinter', 'customtkinter'} & set(sys.modules)))"

    completed = subprocess.run(
        [sys.executable, "-c", code], cwd=PROJECT_ROOT, capture_output=True, text=True, check=True
    )

    assert completed.stdout.strip() == "[]"


# --- Проверка ввода -------------------------------------------------------------------


def test_mock_request_needs_only_config(tmp_path):
    assert validate_request(request(tmp_path)) is None


@pytest.mark.parametrize(
    ("overrides", "message"),
    [
        ({"config_path": Path("нет.yaml")}, "Конфиг не найден"),
        ({"mock": False}, "Выберите файл выгрузки"),
        ({"mock": False, "input_path": Path("нет.xlsx")}, "Файл выгрузки не найден"),
    ],
)
def test_invalid_requests_are_explained(tmp_path, overrides, message):
    assert message in validate_request(request(tmp_path, **overrides))


def test_unsupported_input_format(tmp_path):
    source = tmp_path / "export.pdf"
    source.write_bytes(b"%PDF")

    problem = validate_request(request(tmp_path, mock=False, input_path=source))

    assert "«.pdf» не поддерживается" in problem and ".xlsx" in problem


# --- Рабочий поток ------------------------------------------------------------------------


def test_worker_reports_each_stage_and_success(tmp_path):
    events: queue.Queue = queue.Queue()

    worker = threading.Thread(target=run_generation, args=(request(tmp_path, notes="- Первый вывод"), events))
    worker.start()
    worker.join(timeout=60)

    items = drain(events)
    progress = [item for item in items if isinstance(item, ProgressEvent)]
    assert [item.label for item in progress] == ["Чтение конфига", "Чтение данных", "Агрегация", "Рендеринг"]
    assert [item.fraction for item in progress] == [0.0, 0.25, 0.5, 0.75]
    assert isinstance(items[-1], SuccessEvent)
    result = items[-1].result
    assert result.output_path.parent == (tmp_path / "output").resolve() and result.output_path.exists()
    assert result.has_notes
    assert result.report_month is None  # тестовые данные в историю не попадают
    assert not (tmp_path / "history.db").exists()


def test_worker_turns_errors_into_failure_event(tmp_path):
    events: queue.Queue = queue.Queue()
    broken = tmp_path / "broken.xlsx"
    broken.write_text("не Excel", encoding="utf-8")

    run_generation(request(tmp_path, mock=False, input_path=broken), events)

    [*_, last] = drain(events)
    assert isinstance(last, FailureEvent)
    assert last.title == "Не удалось собрать отчёт" and "не похож на книгу Excel" in last.message


def test_worker_reports_locked_report(tmp_path, monkeypatch):
    def locked(*args, **kwargs):
        raise PermissionError(13, "Permission denied", "report.pptx")

    monkeypatch.setattr("src.gui.tasks.generate_report", locked)
    events: queue.Queue = queue.Queue()

    run_generation(request(tmp_path), events)

    [event] = drain(events)
    assert event.title == "Файл занят" and "PowerPoint" in event.message


def test_worker_survives_unexpected_errors(tmp_path, monkeypatch):
    monkeypatch.setattr("src.gui.tasks.generate_report", lambda *args, **kwargs: 1 / 0)
    events: queue.Queue = queue.Queue()

    run_generation(request(tmp_path), events)

    [event] = drain(events)
    assert event.title == "Непредвиденная ошибка" and "ZeroDivisionError" in event.message


def test_log_handler_forwards_records_to_window():
    events: queue.Queue = queue.Queue()
    logger = logging.getLogger("src.test_gui_tasks")
    handler = QueueLogHandler(events)
    logger.addHandler(handler)
    try:
        logger.warning("Выгрузка охватывает %d дн.", 31)
    finally:
        logger.removeHandler(handler)

    [event] = drain(events)
    assert isinstance(event, LogEvent) and event.level == logging.WARNING
    assert event.message.endswith("WARNING Выгрузка охватывает 31 дн.")


# --- Конфиги и заметки ------------------------------------------------------------------------


def test_config_list_and_description(tmp_path):
    (tmp_path / "b_client.yml").write_text(DEFAULT_CONFIG_PATH.read_text(encoding="utf-8"), encoding="utf-8")
    (tmp_path / "A_client.yaml").write_text("report_metadata: {client_name: Клиент}\nslides: []\n", encoding="utf-8")
    (tmp_path / "notes.md").write_text("не конфиг", encoding="utf-8")

    files = list_config_files(tmp_path)

    assert [path.name for path in files] == ["A_client.yaml", "b_client.yml"]
    valid, text = describe_config(files[1])
    assert valid and text.startswith("ООО Доставка Плюс · период 30 дн.") and "маппинг колонок: 6" in text
    valid, text = describe_config(files[0])
    assert not valid and "slides" in text
    assert list_config_files(tmp_path / "нет") == []


def test_read_notes_file(tmp_path):
    notes = tmp_path / "notes.md"
    notes.write_text("# Итоги\n- Пункт", encoding="utf-8")

    assert read_notes_file(notes) == "# Итоги\n- Пункт"
    with pytest.raises(ValueError, match=".md и .txt"):
        read_notes_file(tmp_path / "notes.docx")


# --- Первый запуск собранного .exe ---------------------------------------------------------------


def test_bundled_configs_are_copied_on_first_run(tmp_path):
    bundled = tmp_path / "bundle" / "config"
    bundled.mkdir(parents=True)
    (bundled / "report_config.yaml").write_text("a: 1", encoding="utf-8")
    (bundled / "readme.txt").write_text("не конфиг", encoding="utf-8")
    user_dir = tmp_path / "app" / "config"

    copied = ensure_user_configs(bundled, user_dir)

    assert copied == [user_dir / "report_config.yaml"]
    assert sorted(path.name for path in user_dir.iterdir()) == ["report_config.yaml"]

    (user_dir / "report_config.yaml").unlink()
    (user_dir / "client_b.yaml").write_text("b: 2", encoding="utf-8")
    assert ensure_user_configs(bundled, user_dir) == []  # у пользователя свои конфиги — не трогаем


def test_configs_are_not_copied_onto_themselves(tmp_path):
    (tmp_path / "report_config.yaml").write_text("a: 1", encoding="utf-8")

    assert ensure_user_configs(tmp_path, tmp_path) == []
