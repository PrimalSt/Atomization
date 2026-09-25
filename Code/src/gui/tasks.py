"""Логика окна без tkinter: конфиги, проверка ввода, сборка отчёта в рабочем потоке.

Рабочий поток не трогает виджеты — tkinter не потокобезопасен. Он кладёт события
в ``queue.Queue``, а окно забирает их таймером в главном потоке.
"""

import logging
import os
import queue
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

from pydantic import ValidationError

from src.config_loader import load_report_config
from src.data_loader import SUPPORTED_SUFFIXES
from src.direct_client import DirectApiError
from src.paths import CONFIG_SUFFIXES
from src.pipeline import NOTES_SUFFIXES, ReportResult, generate_report, read_text_file

logger = logging.getLogger(__name__)

# Шаги пайплайна → подписи этапов в строке статуса.
STAGE_LABELS: dict[str, str] = {
    "Конфиг": "Чтение конфига",
    "Данные": "Чтение данных",
    "Агрегация": "Агрегация",
    "Рендеринг": "Рендеринг",
}


# ---------------------------------------------------------------------------
# События рабочего потока
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class ProgressEvent:
    step: int
    total: int
    name: str  # шаг пайплайна: «Конфиг», «Данные», «Агрегация», «Рендеринг»

    @property
    def label(self) -> str:
        return STAGE_LABELS.get(self.name, self.name)

    @property
    def fraction(self) -> float:
        """Доля выполненного к началу шага: 0 перед первым, 0,75 перед последним из четырёх."""
        return (self.step - 1) / self.total if self.total else 0.0


@dataclass(frozen=True, slots=True)
class LogEvent:
    level: int
    message: str


@dataclass(frozen=True, slots=True)
class SuccessEvent:
    result: ReportResult


@dataclass(frozen=True, slots=True)
class FailureEvent:
    title: str
    message: str


Event = ProgressEvent | LogEvent | SuccessEvent | FailureEvent


class QueueLogHandler(logging.Handler):
    """Пересылает записи лога в очередь окна: предупреждения пайплайна видны в журнале GUI."""

    def __init__(self, events: "queue.Queue[Event]", level: int = logging.INFO) -> None:
        super().__init__(level)
        self.events = events
        self.setFormatter(logging.Formatter("%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S"))

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.events.put(LogEvent(record.levelno, self.format(record)))
        except Exception:  # noqa: BLE001 — сбой журнала не должен ронять сборку отчёта
            self.handleError(record)


# ---------------------------------------------------------------------------
# Запрос на сборку
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class GenerationRequest:
    config_path: Path
    input_path: Path | None
    mock: bool
    notes: str
    output_dir: Path
    history_db: Path | None  # None — не вести историю и не считать MoM


def validate_request(request: GenerationRequest) -> str | None:
    """Текст ошибки для пользователя или None, если с запросом можно запускать сборку."""
    if not request.config_path.is_file():
        return f"Конфиг не найден: {request.config_path}"
    if request.mock:
        return None
    if request.input_path is None or not str(request.input_path).strip():
        return "Выберите файл выгрузки или включите «Использовать тестовые данные (Mock)»."
    if not request.input_path.is_file():
        return f"Файл выгрузки не найден: {request.input_path}"
    if request.input_path.suffix.lower() not in SUPPORTED_SUFFIXES:
        formats = ", ".join(sorted(SUPPORTED_SUFFIXES))
        return f"Формат «{request.input_path.suffix or 'без расширения'}» не поддерживается. Подходят: {formats}."
    return None


def run_generation(request: GenerationRequest, events: "queue.Queue[Event]") -> None:
    """Тело рабочего потока: собирает отчёт и сообщает о ходе и итоге через очередь.

    Исключения не выпускает: любой исход — событие ``SuccessEvent`` или ``FailureEvent``.
    """

    def report_progress(step: int, total: int, name: str) -> None:
        events.put(ProgressEvent(step, total, name))

    try:
        result = generate_report(
            request.config_path,
            input_path=None if request.mock else request.input_path,
            mock=request.mock,
            notes=request.notes if request.notes.strip() else None,
            output=request.output_dir,
            history_db=None if request.mock else request.history_db,
            on_progress=report_progress,
        )
    except PermissionError as exc:
        events.put(
            FailureEvent(
                "Файл занят",
                f"Нет доступа к файлу {exc.filename or exc}.\n"
                "Если отчёт с таким именем открыт в PowerPoint, закройте его и повторите.",
            )
        )
    except (OSError, ValueError, DirectApiError) as exc:  # pydantic.ValidationError — подкласс ValueError
        events.put(FailureEvent("Не удалось собрать отчёт", str(exc)))
    except Exception as exc:  # noqa: BLE001 — окно должно показать ошибку, а не зависнуть в «Рендеринге»
        logger.exception("Непредвиденная ошибка при сборке отчёта")
        events.put(FailureEvent("Непредвиденная ошибка", f"{type(exc).__name__}: {exc}"))
    else:
        events.put(SuccessEvent(result))


# ---------------------------------------------------------------------------
# Конфиги, заметки, файлы
# ---------------------------------------------------------------------------


def list_config_files(config_dir: Path) -> list[Path]:
    """YAML-конфиги клиентов из папки, по имени файла."""
    if not config_dir.is_dir():
        return []
    files = (path for path in config_dir.iterdir() if path.is_file() and path.suffix.lower() in CONFIG_SUFFIXES)
    return sorted(files, key=lambda path: path.name.casefold())


def describe_config(path: Path) -> tuple[bool, str]:
    """(валиден ли конфиг, строка для окна): клиент, период и слайды — или причина ошибки."""
    try:
        config = load_report_config(path)
    except ValidationError as exc:
        first = exc.errors()[0]
        location = ".".join(str(part) for part in first["loc"])
        return False, f"Конфиг не прошёл валидацию ({exc.error_count()} ош.): {location} — {first['msg']}"
    except (OSError, ValueError) as exc:
        return False, f"Не удалось прочитать конфиг: {exc}"
    metadata = config.report_metadata
    mapping = f" · маппинг колонок: {len(config.column_mapping)}" if config.column_mapping else ""
    return True, (
        f"{metadata.client_name} · период {metadata.period_days} дн. · слайдов: {len(config.active_slides)}{mapping}"
    )


def read_notes_file(path: Path) -> str:
    """Текст заметок из .md / .txt (UTF-8 или cp1251).

    Raises:
        ValueError: расширение не .txt / .md.
        OSError: файл не читается.
    """
    if path.suffix.lower() not in NOTES_SUFFIXES:
        raise ValueError(f"Заметки читаются из файлов .md и .txt, получено «{path.name}»")
    return read_text_file(path)


def open_path(path: Path) -> None:
    """Открывает файл или папку в программе по умолчанию (отчёт — в PowerPoint).

    Raises:
        OSError: файла нет или для него не назначена программа.
    """
    if not path.exists():
        raise FileNotFoundError(f"Не найден: {path}")
    if sys.platform == "win32":
        os.startfile(path)  # type: ignore[attr-defined]  # есть только в Windows
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])
