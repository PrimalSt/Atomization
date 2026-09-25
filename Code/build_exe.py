"""Сборка десктопного приложения в исполняемый файл через PyInstaller.

    python build_exe.py              # папка dist/AutoReport/ с AutoReport.exe — быстрый запуск
    python build_exe.py --onefile    # один файл dist/AutoReport.exe — удобно пересылать
    python build_exe.py --icon app.ico

Собирать нужно на Windows: PyInstaller не делает кросс-компиляцию, на Linux и macOS
получится исполняемый файл для этой ОС. Нужны зависимости проекта и PyInstaller:
``pip install -r requirements-dev.txt``.

В бандл попадают конфиги из config/ (при первом запуске приложение копирует их в папку
config/ рядом с .exe — там их можно править и добавлять новых клиентов), темы и шрифты
customtkinter и шаблон презентации python-pptx. Отчёты (output/) и история
(local_storage/history.db) создаются рядом с .exe, поэтому распакуйте сборку в папку,
куда у пользователя есть права на запись (не в Program Files).
"""

import argparse
import importlib.util
import os
import shutil
import sys
from collections.abc import Sequence
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
APP_NAME = "AutoReport"
ENTRY_POINT = PROJECT_ROOT / "app_gui.py"
CONFIG_DIR = PROJECT_ROOT / "config"
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"

# Пакеты, чьи файлы данных PyInstaller сам не видит: темы/шрифты customtkinter и default.pptx.
DATA_PACKAGES = ("customtkinter", "pptx")
REQUIRED_MODULES = {
    "PyInstaller": "pip install -r requirements-dev.txt",
    "tkinter": "установите Python с компонентом «tcl/tk and IDLE» (python.org)",
    "customtkinter": "pip install -r requirements.txt",
    "openpyxl": "pip install -r requirements.txt",
    "xlrd": "pip install -r requirements.txt",
    "pptx": "pip install -r requirements.txt",
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Сборка AutoReport в исполняемый файл (PyInstaller, --windowed).")
    parser.add_argument("--onefile", action="store_true", help="один .exe вместо папки (запускается медленнее)")
    parser.add_argument("--icon", type=Path, metavar="ICO", help="иконка приложения (.ico)")
    parser.add_argument("--name", default=APP_NAME, help=f"имя исполняемого файла (по умолчанию {APP_NAME})")
    return parser


def check_environment() -> list[str]:
    """Чего не хватает для сборки: по строке с подсказкой на каждый модуль."""
    problems = []
    for module, hint in REQUIRED_MODULES.items():
        if importlib.util.find_spec(module) is None:
            problems.append(f"нет модуля {module}: {hint}")
    if not ENTRY_POINT.is_file():
        problems.append(f"нет точки входа {ENTRY_POINT}")
    if not any(CONFIG_DIR.glob("*.yaml")) and not any(CONFIG_DIR.glob("*.yml")):
        problems.append(f"в {CONFIG_DIR} нет ни одного конфига .yaml")
    return problems


def pyinstaller_args(*, name: str, onefile: bool, icon: Path | None) -> list[str]:
    """Аргументы командной строки PyInstaller. Пути абсолютные: spec-файл лежит в build/."""
    args = [
        str(ENTRY_POINT),
        "--name", name,
        "--windowed",  # без чёрного окна консоли
        "--onefile" if onefile else "--onedir",
        "--noconfirm",
        "--clean",
        "--distpath", str(DIST_DIR),
        "--workpath", str(BUILD_DIR),
        "--specpath", str(BUILD_DIR),
        "--paths", str(PROJECT_ROOT),
        "--add-data", f"{CONFIG_DIR}{os.pathsep}config",
    ]  # fmt: skip
    for package in DATA_PACKAGES:
        args += ["--collect-data", package]
    # python-pptx открывает шаблоны по пути pptx/oxml/../templates/*.xml. В бандле модули лежат
    # в архиве PYZ, папки pptx/oxml на диске нет, и в Linux/macOS такой путь не открывается
    # (Windows сокращает «..» без проверки папки). Файл в pptx/oxml создаёт эту папку.
    args += ["--add-data", f"{_module_file('pptx.oxml')}{os.pathsep}pptx/oxml"]
    if icon is not None:
        args += ["--icon", str(icon.resolve())]
    return args


def _module_file(module: str) -> str:
    spec = importlib.util.find_spec(module)
    if spec is None or spec.origin is None:
        raise RuntimeError(f"Не найден модуль {module}")
    return spec.origin


def app_folder(*, name: str, onefile: bool) -> Path:
    """Папка, где лежит исполняемый файл: рядом с ним приложение хранит конфиги, отчёты и историю."""
    return DIST_DIR if onefile else DIST_DIR / name


def executable_path(*, name: str, onefile: bool) -> Path:
    suffix = ".exe" if sys.platform == "win32" else ""
    return app_folder(name=name, onefile=onefile) / f"{name}{suffix}"


def copy_user_configs(target_dir: Path) -> list[Path]:
    """Кладёт конфиги рядом с .exe, не затирая уже отредактированные при пересборке."""
    target_dir.mkdir(parents=True, exist_ok=True)
    copied = []
    for source in sorted(CONFIG_DIR.iterdir()):
        target = target_dir / source.name
        if source.is_file() and source.suffix.lower() in {".yaml", ".yml"} and not target.exists():
            copied.append(Path(shutil.copy2(source, target)))
    return copied


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.icon is not None and not args.icon.is_file():
        print(f"Иконка не найдена: {args.icon}", file=sys.stderr)
        return 1
    problems = check_environment()
    if problems:
        print("Сборка невозможна:\n  " + "\n  ".join(problems), file=sys.stderr)
        return 1
    if sys.platform != "win32":
        print(f"Внимание: сборка на {sys.platform} даёт исполняемый файл для этой ОС, а не .exe для Windows.")

    import PyInstaller.__main__  # импорт после проверки окружения: без PyInstaller — понятное сообщение

    PyInstaller.__main__.run(pyinstaller_args(name=args.name, onefile=args.onefile, icon=args.icon))

    folder = app_folder(name=args.name, onefile=args.onefile)
    copy_user_configs(folder / "config")
    executable = executable_path(name=args.name, onefile=args.onefile)
    if not executable.is_file():
        print(f"PyInstaller завершился, но {executable} не найден — смотрите лог выше", file=sys.stderr)
        return 1
    print(f"\nГотово: {executable}")
    print(f"Конфиги клиентов: {folder / 'config'} — правьте их там или добавляйте новые .yaml")
    print(f"Для передачи другому компьютеру заархивируйте папку {folder} целиком.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
