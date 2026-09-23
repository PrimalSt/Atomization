"""Командная строка генератора отчётов по контекстной рекламе.

    python main.py --mock
    python main.py -i export.csv -n "CPA снизился на 12 %"
    python main.py -i export.json --notes-file notes.md -o reports/
"""

import argparse
import logging
import sys
from collections.abc import Sequence
from pathlib import Path

from src.config_loader import DEFAULT_CONFIG_PATH
from src.console import configure_logging, paint, supports_color, use_utf8_console
from src.pipeline import DEFAULT_MOCK_CAMPAIGNS, ReportResult, generate_report
from src.rendering.builder import OUTPUT_DIR

logger = logging.getLogger("report.cli")

_EXAMPLES = """примеры:
  python main.py --mock
  python main.py --mock --seed 7 -o output/demo.pptx
  python main.py -i export.csv -n "CPA снизился на 12 %"
  python main.py -i export.csv -n "- Отключили РСЯ-площадки с CPA выше 3 000 ₽\\n- Тестируем новые объявления"
  python main.py -i export.json --notes-file notes.md -o reports/

В --notes последовательность \\n означает перенос строки. Для развёрнутых выводов
удобнее --notes-file: файл .txt или .md с заголовками (#), абзацами и списками (-, 1.).
"""


def _positive_int(value: str) -> int:
    try:
        number = int(value)
    except ValueError:
        raise argparse.ArgumentTypeError(f"ожидается целое число, получено «{value}»") from None
    if number < 1:
        raise argparse.ArgumentTypeError(f"ожидается число от 1, получено {number}")
    return number


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="main.py",
        description="Собирает PowerPoint-отчёт по контекстной рекламе: конфиг → данные → агрегация → презентация.",
        epilog=_EXAMPLES,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    source = parser.add_mutually_exclusive_group()
    source.add_argument("-i", "--input", type=Path, metavar="ФАЙЛ", help="выгрузка кабинета: .csv, .tsv или .json")
    source.add_argument("--mock", action="store_true", help="сгенерировать синтетические данные за период из конфига")
    parser.add_argument(
        "-c", "--config", type=Path, default=DEFAULT_CONFIG_PATH, metavar="YAML",
        help="конфиг отчёта (по умолчанию config/report_config.yaml)",
    )  # fmt: skip
    parser.add_argument(
        "-o", "--output", type=Path, default=OUTPUT_DIR, metavar="ПУТЬ",
        help="папка или файл .pptx (по умолчанию output/, имя файла — из клиента и периода)",
    )  # fmt: skip

    notes = parser.add_mutually_exclusive_group()
    notes.add_argument("-n", "--notes", metavar="ТЕКСТ", help="выводы специалиста для слайда notes_slide")
    notes.add_argument("--notes-file", type=Path, metavar="ФАЙЛ", help="файл .txt или .md с выводами специалиста")

    mock = parser.add_argument_group("синтетические данные (только с --mock)")
    mock.add_argument(
        "--campaigns", type=_positive_int, metavar="N", help=f"число кампаний (по умолчанию {DEFAULT_MOCK_CAMPAIGNS})"
    )
    mock.add_argument("--seed", type=int, metavar="N", help="зерно генератора — одинаковые данные при каждом запуске")

    parser.add_argument("-v", "--verbose", action="store_true", help="подробный лог (DEBUG)")
    parser.add_argument("--no-color", action="store_true", help="без цветов в консоли")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """Точка входа CLI. Возвращает код выхода: 0 — отчёт собран, 1 — ошибка входных данных."""
    use_utf8_console()
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.input is None and not args.mock:
        parser.error("укажите источник данных: -i/--input ФАЙЛ или --mock")

    configure_logging(verbose=args.verbose, color=not args.no_color)
    if args.input is not None and (args.campaigns is not None or args.seed is not None):
        logger.warning("--campaigns и --seed действуют только вместе с --mock — параметры не применены")
    notes = args.notes.replace("\\n", "\n") if args.notes is not None else args.notes_file

    try:
        result = generate_report(
            args.config,
            input_path=args.input,
            mock=args.mock,
            notes=notes,
            output=args.output,
            mock_campaigns=args.campaigns or DEFAULT_MOCK_CAMPAIGNS,
            mock_seed=args.seed,
        )
    except PermissionError as exc:
        logger.error("Нет доступа к файлу %s — если отчёт открыт в PowerPoint, закройте его", exc.filename or exc)
        return 1
    except (OSError, ValueError) as exc:  # pydantic.ValidationError — подкласс ValueError
        logger.error("%s", exc)
        return 1
    except KeyboardInterrupt:
        logger.error("Прервано пользователем")
        return 130

    _print_summary(result, color=not args.no_color and supports_color(sys.stdout))
    return 0


def _print_summary(result: ReportResult, *, color: bool) -> None:
    """Итог — в stdout (лог шагов идёт в stderr), чтобы путь к отчёту можно было забрать скриптом."""
    date_from, date_to = result.period
    notes = "выводы специалиста добавлены" if result.has_notes else "на слайде выводов — заготовка"
    print(f"{paint('Отчёт готов:', 'green', 'bold', enabled=color)} {result.output_path}")
    print(
        f"  период {date_from:%d.%m.%Y} – {date_to:%d.%m.%Y} · записей: {result.records_count} · "
        f"кампаний: {result.campaigns_count} · слайдов: {result.slides_count} · {notes}"
    )
    print(paint(f"  время: {result.timing_summary()}", "dim", enabled=color))


if __name__ == "__main__":
    sys.exit(main())
