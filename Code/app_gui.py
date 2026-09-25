"""Десктопное приложение генератора отчётов: python app_gui.py

Окно выбирает конфиг клиента из config/, файл выгрузки (.xlsx, .xls, .csv, .tsv,
.json) или тестовые данные, принимает выводы специалиста и собирает презентацию
в output/. Консольный вариант — main.py.
"""

import sys


def main() -> int:
    try:
        from src.gui.app import run
    except ModuleNotFoundError as exc:
        if exc.name not in {"tkinter", "_tkinter", "customtkinter"}:
            raise
        hint = (
            "установите зависимости: pip install -r requirements.txt"
            if exc.name == "customtkinter"
            else "установите Python с компонентом «tcl/tk and IDLE» (python.org)"
        )
        print(f"Графический интерфейс недоступен: нет модуля {exc.name} — {hint}", file=sys.stderr)
        return 1
    return run()


if __name__ == "__main__":
    sys.exit(main())
