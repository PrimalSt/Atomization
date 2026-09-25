"""Пути приложения: запуск из исходников и из собранного PyInstaller'ом .exe.

Из исходников всё лежит в корне проекта, как и раньше. В собранном приложении
файлы поставки (конфиги по умолчанию) распакованы во временную или служебную папку
бандла, а данные пользователя — конфиги клиентов, отчёты, история, ``.env`` —
должны переживать перезапуск, поэтому живут рядом с исполняемым файлом.
"""

import shutil
import sys
from pathlib import Path

FROZEN = bool(getattr(sys, "frozen", False))

_SOURCE_ROOT = Path(__file__).resolve().parent.parent

# Файлы из поставки: в .exe — распакованный бандл (sys._MEIPASS), из исходников — корень проекта.
BUNDLE_ROOT = Path(getattr(sys, "_MEIPASS", _SOURCE_ROOT))
# Данные пользователя: в .exe — папка с исполняемым файлом, из исходников — корень проекта.
APP_ROOT = Path(sys.executable).resolve().parent if FROZEN else _SOURCE_ROOT

CONFIG_DIR = APP_ROOT / "config"
OUTPUT_DIR = APP_ROOT / "output"
STORAGE_DIR = APP_ROOT / "local_storage"
CONFIG_SUFFIXES = frozenset({".yaml", ".yml"})


def ensure_user_configs(bundled_dir: Path = BUNDLE_ROOT / "config", config_dir: Path = CONFIG_DIR) -> list[Path]:
    """Первый запуск собранного приложения: копирует конфиги из бандла в папку пользователя.

    Копирует, только если в ``config_dir`` ещё нет ни одного YAML: так удалённый
    пользователем конфиг не возвращается при каждом запуске. Из исходников обе папки
    совпадают, и функция ничего не делает.

    Returns:
        Скопированные файлы.
    """
    if not bundled_dir.is_dir() or bundled_dir.resolve() == config_dir.resolve():
        return []
    if config_dir.is_dir() and any(path.suffix.lower() in CONFIG_SUFFIXES for path in config_dir.iterdir()):
        return []
    config_dir.mkdir(parents=True, exist_ok=True)
    copied = []
    for source in sorted(bundled_dir.iterdir()):
        if source.is_file() and source.suffix.lower() in CONFIG_SUFFIXES:
            copied.append(Path(shutil.copy2(source, config_dir / source.name)))
    return copied
