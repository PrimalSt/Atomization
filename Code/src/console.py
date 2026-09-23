"""Консольный вывод: UTF-8 на русской Windows, цвета ANSI и формат логов."""

import codecs
import io
import logging
import os
import sys
from typing import TextIO

# Пакет сервиса, CLI (main.py) и модуль src, запущенный как скрипт (python -m src.data_loader).
LOGGER_NAMES = ("src", "report", "__main__")

_ANSI = {"bold": "1", "dim": "2", "red": "31", "green": "32", "yellow": "33", "cyan": "36"}
_LEVEL_STYLES = {
    logging.DEBUG: ("dim",),
    logging.INFO: ("cyan",),
    logging.WARNING: ("yellow", "bold"),
    logging.ERROR: ("red", "bold"),
    logging.CRITICAL: ("red", "bold"),
}
_ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004


def use_utf8_console() -> None:
    """Переключает stdout и stderr на UTF-8, если они не в нём.

    При перенаправлении в файл или pipe Windows по умолчанию пишет в cp1251,
    где нет «₽» и типографских символов. На вывод в саму консоль это не влияет:
    туда Python пишет через WriteConsoleW.
    """
    for stream in (sys.stdout, sys.stderr):
        if isinstance(stream, io.TextIOWrapper) and codecs.lookup(stream.encoding).name != "utf-8":
            stream.reconfigure(encoding="utf-8")


def supports_color(stream: TextIO) -> bool:
    """Цвета — только в интерактивном терминале и если пользователь не задал NO_COLOR."""
    if os.environ.get("NO_COLOR") or not hasattr(stream, "isatty") or not stream.isatty():
        return False
    return _enable_windows_ansi(stream) if sys.platform == "win32" else True


def _enable_windows_ansi(stream: TextIO) -> bool:
    """Включает обработку ANSI-последовательностей в консоли Windows 10+."""
    import ctypes
    import msvcrt
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.GetConsoleMode.argtypes = (wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD))
    kernel32.SetConsoleMode.argtypes = (wintypes.HANDLE, wintypes.DWORD)
    try:
        handle = msvcrt.get_osfhandle(stream.fileno())
    except (OSError, ValueError, io.UnsupportedOperation):
        return False
    mode = wintypes.DWORD()
    if not kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
        return False
    return bool(kernel32.SetConsoleMode(handle, mode.value | _ENABLE_VIRTUAL_TERMINAL_PROCESSING))


def paint(text: str, *styles: str, enabled: bool = True) -> str:
    if not enabled or not styles:
        return text
    codes = ";".join(_ANSI[style] for style in styles)
    return f"\033[{codes}m{text}\033[0m"


class ConsoleFormatter(logging.Formatter):
    """«12:30:05 INFO    сообщение»; уровень подсвечивается цветом."""

    def __init__(self, *, color: bool) -> None:
        super().__init__()
        self.color = color

    def format(self, record: logging.LogRecord) -> str:
        message = record.getMessage()
        if record.exc_info:
            message = f"{message}\n{self.formatException(record.exc_info)}"
        time = paint(self.formatTime(record, "%H:%M:%S"), "dim", enabled=self.color)
        level = paint(f"{record.levelname:<7}", *_LEVEL_STYLES.get(record.levelno, ()), enabled=self.color)
        return f"{time} {level} {message}"


def configure_logging(*, verbose: bool = False, color: bool = True) -> bool:
    """Направляет логи сервиса в stderr в консольном формате.

    Повторный вызов заменяет обработчик, а не добавляет второй. Возвращает,
    включены ли цвета (их нет при перенаправлении вывода или с NO_COLOR).
    """
    use_color = color and supports_color(sys.stderr)
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(ConsoleFormatter(color=use_color))
    handler.set_name("report-console")
    for name in LOGGER_NAMES:
        logger = logging.getLogger(name)
        for existing in [h for h in logger.handlers if h.get_name() == "report-console"]:
            logger.removeHandler(existing)
        logger.addHandler(handler)
        logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    return use_color
