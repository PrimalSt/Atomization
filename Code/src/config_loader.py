"""Загрузка YAML-конфигурации отчёта с валидацией через ReportConfig."""

import logging
from pathlib import Path

import yaml

from src.paths import APP_ROOT, CONFIG_DIR
from src.schemas import ReportConfig

PROJECT_ROOT = APP_ROOT  # корень проекта; в собранном .exe — папка с исполняемым файлом
DEFAULT_CONFIG_PATH = CONFIG_DIR / "report_config.yaml"

logger = logging.getLogger(__name__)


def load_report_config(path: str | Path = DEFAULT_CONFIG_PATH) -> ReportConfig:
    """Читает YAML-конфиг и валидирует его.

    Raises:
        FileNotFoundError: файла нет.
        ValueError: YAML синтаксически некорректен или его корень — не словарь.
        pydantic.ValidationError: структура или значения не соответствуют схеме.
    """
    config_path = Path(path)
    with config_path.open(encoding="utf-8") as config_file:
        try:
            raw = yaml.safe_load(config_file)
        except yaml.YAMLError as exc:
            raise ValueError(f"Некорректный YAML в {config_path}: {exc}") from exc
    if not isinstance(raw, dict):
        raise ValueError(f"Корень {config_path} должен быть словарём, получено: {type(raw).__name__}")
    config = ReportConfig.model_validate(raw)
    logger.debug("Конфиг %s валиден: активных слайдов %d", config_path, len(config.active_slides))
    return config
