"""Окно десктопного приложения: конфиг клиента → выгрузка → выводы → презентация.

Запуск: ``python app_gui.py``. Сборка отчёта идёт в отдельном потоке
(``threading.Thread``), окно получает от него события через очередь и остаётся
отзывчивым всё время сборки.
"""

import logging
import queue
import threading
import tkinter as tk
from collections.abc import Sequence
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from src.gui.tasks import (
    Event,
    FailureEvent,
    GenerationRequest,
    LogEvent,
    ProgressEvent,
    QueueLogHandler,
    SuccessEvent,
    describe_config,
    list_config_files,
    open_path,
    read_notes_file,
    run_generation,
    validate_request,
)
from src.paths import CONFIG_DIR, OUTPUT_DIR, ensure_user_configs
from src.pipeline import ReportResult
from src.storage import DEFAULT_DB_PATH

logger = logging.getLogger(__name__)

APP_TITLE = "Генератор отчётов по контекстной рекламе"
_POLL_MS = 100  # как часто окно забирает события рабочего потока
_MAX_EVENTS_PER_POLL = 200
_MAX_LOG_LINES = 500
_PAD = 16

_APPEARANCE_MODES = {"Системная": "system", "Светлая": "light", "Тёмная": "dark"}
_DATA_FILETYPES = [
    ("Выгрузки кабинета", "*.xlsx *.xlsm *.xls *.csv *.tsv *.json"),
    ("Excel", "*.xlsx *.xlsm *.xls"),
    ("CSV / TSV", "*.csv *.tsv"),
    ("JSON", "*.json"),
    ("Все файлы", "*.*"),
]
_NOTES_FILETYPES = [("Заметки", "*.md *.markdown *.txt"), ("Все файлы", "*.*")]
_NOTES_HINT = (
    "Markdown: # заголовок · - пункт списка · 1. нумерованный пункт · **полужирный** · пустая строка — новый абзац"
)


class ReportApp(ctk.CTk):
    """Главное окно генератора отчётов."""

    def __init__(
        self,
        *,
        config_dir: Path = CONFIG_DIR,
        output_dir: Path = OUTPUT_DIR,
        history_db: Path = DEFAULT_DB_PATH,
        notices: Sequence[str] = (),
    ) -> None:
        """``notices`` — сообщения, возникшие до открытия окна; они попадут в журнал."""
        super().__init__()
        self.config_dir = config_dir
        self.output_dir = output_dir
        self.history_db = history_db

        self._events: queue.Queue[Event] = queue.Queue()
        self._worker: threading.Thread | None = None
        self._last_report: Path | None = None
        self._configs: dict[str, Path] = {}
        self._last_data_dir = Path.home()
        self._last_notes_dir = Path.home()

        self._log_handler = QueueLogHandler(self._events)
        self._source_logger = logging.getLogger("src")
        self._source_logger.addHandler(self._log_handler)
        if self._source_logger.level == logging.NOTSET or self._source_logger.level > logging.INFO:
            self._source_logger.setLevel(logging.INFO)

        self.title(APP_TITLE)
        height = min(900, max(640, self.winfo_screenheight() - 100))
        self.geometry(f"860x{height}")
        self.minsize(720, 640)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.mock_var = tk.BooleanVar(value=False)
        self.history_var = tk.BooleanVar(value=True)
        self.input_var = tk.StringVar(value="")

        self._build_layout()
        for notice in notices:
            self._append_log(logging.INFO, notice)
        self.refresh_configs()
        self._sync_source_state()
        self._poll_job = self.after(_POLL_MS, self._poll_events)

    # --- Разметка -------------------------------------------------------------

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=3)  # выводы специалиста
        self.grid_rowconfigure(7, weight=1)  # журнал

        self._build_header(row=0)
        self._build_config_section(row=1)
        self._build_data_section(row=2)
        self._build_notes_section(row=3)

        self.generate_button = ctk.CTkButton(
            self,
            text="Сгенерировать презентацию (.pptx)",
            height=44,
            font=ctk.CTkFont(size=16, weight="bold"),
            command=self.start_generation,
        )
        self.generate_button.grid(row=4, column=0, sticky="ew", padx=_PAD, pady=(4, 8))

        self._build_progress(row=5)
        self._build_result_actions(row=6)
        self._build_log(row=7)

    def _section(self, row: int, title: str) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self)
        frame.grid(row=row, column=0, sticky="nsew", padx=_PAD, pady=(0, 8))
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=14, weight="bold"), anchor="w").grid(
            row=0, column=0, columnspan=3, sticky="ew", padx=12, pady=(8, 4)
        )
        return frame

    def _build_header(self, row: int) -> None:
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=row, column=0, sticky="ew", padx=_PAD, pady=(_PAD, 8))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text=APP_TITLE, font=ctk.CTkFont(size=20, weight="bold"), anchor="w").grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkLabel(
            header, text=f"Отчёты сохраняются в {self.output_dir}", text_color=("gray35", "gray65"),
            anchor="w", justify="left", wraplength=560,
        ).grid(row=1, column=0, sticky="w")  # fmt: skip
        self.appearance_switch = ctk.CTkSegmentedButton(
            header, values=list(_APPEARANCE_MODES), command=self._change_appearance
        )
        self.appearance_switch.set("Системная")
        self.appearance_switch.grid(row=0, column=1, rowspan=2, sticky="e")

    def _build_config_section(self, row: int) -> None:
        frame = self._section(row, "1. Клиент и конфиг отчёта")
        self.config_box = ctk.CTkComboBox(frame, values=[], state="readonly", command=self._on_config_selected)
        self.config_box.grid(row=1, column=0, sticky="ew", padx=(12, 8), pady=4)
        ctk.CTkButton(frame, text="Обновить список", width=140, command=self.refresh_configs).grid(
            row=1, column=1, padx=(0, 12), pady=4
        )
        self.config_info = ctk.CTkLabel(frame, text="", anchor="w", justify="left", wraplength=760)
        self.config_info.grid(row=2, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 10))

    def _build_data_section(self, row: int) -> None:
        frame = self._section(row, "2. Выгрузка рекламного кабинета")
        self.input_entry = ctk.CTkEntry(
            frame, textvariable=self.input_var, placeholder_text="Файл .xlsx, .xls, .csv, .tsv или .json"
        )
        self.input_entry.grid(row=1, column=0, sticky="ew", padx=(12, 8), pady=4)
        self.browse_button = ctk.CTkButton(frame, text="Обзор...", width=140, command=self.browse_input)
        self.browse_button.grid(row=1, column=1, padx=(0, 12), pady=4)
        self.mock_checkbox = ctk.CTkCheckBox(
            frame,
            text="Использовать тестовые данные (Mock)",
            variable=self.mock_var,
            command=self._sync_source_state,
        )
        self.mock_checkbox.grid(row=2, column=0, columnspan=2, sticky="w", padx=12, pady=(6, 2))
        self.history_checkbox = ctk.CTkCheckBox(
            frame,
            text="Сохранять итоги месяца в историю и показывать динамику к прошлому месяцу (MoM)",
            variable=self.history_var,
        )
        self.history_checkbox.grid(row=3, column=0, columnspan=2, sticky="w", padx=12, pady=(2, 10))

    def _build_notes_section(self, row: int) -> None:
        frame = self._section(row, "3. Выводы специалиста")
        frame.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(frame, text=_NOTES_HINT, text_color=("gray35", "gray65"), anchor="w", justify="left").grid(
            row=1, column=0, columnspan=3, sticky="ew", padx=12
        )
        self.notes_box = ctk.CTkTextbox(frame, height=140, wrap="word", undo=True)
        self.notes_box.grid(row=2, column=0, columnspan=3, sticky="nsew", padx=12, pady=4)
        buttons = ctk.CTkFrame(frame, fg_color="transparent")
        buttons.grid(row=3, column=0, columnspan=3, sticky="w", padx=12, pady=(2, 10))
        ctk.CTkButton(buttons, text="Загрузить заметки из файла...", command=self.load_notes).pack(side="left")
        ctk.CTkButton(
            buttons, text="Очистить", width=100, fg_color="transparent", border_width=1,
            text_color=("gray10", "gray90"), command=self.clear_notes,
        ).pack(side="left", padx=(8, 0))  # fmt: skip

    def _build_progress(self, row: int) -> None:
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid(row=row, column=0, sticky="ew", padx=_PAD)
        frame.grid_columnconfigure(0, weight=1)
        self.progress = ctk.CTkProgressBar(frame, mode="determinate")
        self.progress.set(0)
        self.progress.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        self.status_label = ctk.CTkLabel(frame, text="Готово к работе", anchor="w")
        self.status_label.grid(row=1, column=0, sticky="ew")

    def _build_result_actions(self, row: int) -> None:
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid(row=row, column=0, sticky="ew", padx=_PAD, pady=(4, 8))
        self.open_report_button = ctk.CTkButton(frame, text="Открыть отчёт", state="disabled", command=self.open_report)
        self.open_report_button.pack(side="left")
        self.open_folder_button = ctk.CTkButton(frame, text="Открыть папку", command=self.open_folder)
        self.open_folder_button.pack(side="left", padx=(8, 0))

    def _build_log(self, row: int) -> None:
        frame = self._section(row, "Журнал")
        frame.grid_rowconfigure(1, weight=1)
        self.log_box = ctk.CTkTextbox(frame, height=90, wrap="word", font=ctk.CTkFont(family="Consolas", size=12))
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 10))
        self.log_box.configure(state="disabled")

    # --- Конфиги ----------------------------------------------------------------

    def refresh_configs(self) -> None:
        """Перечитывает папку config/: новые конфиги клиентов появляются без перезапуска."""
        current = self.config_box.get()
        self._configs = {path.name: path for path in list_config_files(self.config_dir)}
        names = list(self._configs)
        self.config_box.configure(values=names)
        if not names:
            self.config_box.set("")
            self.config_info.configure(
                text=f"В папке {self.config_dir} нет конфигов (.yaml). "
                "Добавьте конфиг клиента и нажмите «Обновить список».",
                text_color=("#B42318", "#F97066"),
            )
            return
        self.config_box.set(current if current in self._configs else names[0])
        self._on_config_selected(self.config_box.get())

    def _on_config_selected(self, name: str) -> None:
        path = self._configs.get(name)
        if path is None:
            return
        valid, text = describe_config(path)
        self.config_info.configure(text=text, text_color=("gray20", "gray80") if valid else ("#B42318", "#F97066"))

    # --- Данные и заметки -----------------------------------------------------------

    def browse_input(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self, title="Выгрузка рекламного кабинета", initialdir=self._last_data_dir, filetypes=_DATA_FILETYPES
        )
        if filename:
            path = Path(filename)
            self._last_data_dir = path.parent
            self.input_var.set(str(path))

    def _sync_source_state(self) -> None:
        """Mock отключает выбор файла и историю: синтетические данные в историю не пишутся."""
        mock = self.mock_var.get()
        state = "disabled" if mock else "normal"
        self.input_entry.configure(state=state)
        self.browse_button.configure(state=state)
        self.history_checkbox.configure(state=state)

    def load_notes(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self, title="Выводы специалиста", initialdir=self._last_notes_dir, filetypes=_NOTES_FILETYPES
        )
        if not filename:
            return
        path = Path(filename)
        self._last_notes_dir = path.parent
        try:
            text = read_notes_file(path)
        except (OSError, ValueError) as exc:
            messagebox.showerror("Не удалось загрузить заметки", str(exc), parent=self)
            return
        if self.notes_text().strip() and not messagebox.askyesno(
            "Заменить заметки?", "В поле уже есть текст. Заменить его содержимым файла?", parent=self
        ):
            return
        self.notes_box.delete("1.0", "end")
        self.notes_box.insert("1.0", text)

    def clear_notes(self) -> None:
        self.notes_box.delete("1.0", "end")

    def notes_text(self) -> str:
        return self.notes_box.get("1.0", "end-1c")

    # --- Сборка ---------------------------------------------------------------

    def build_request(self) -> GenerationRequest:
        config_path = self._configs.get(self.config_box.get(), self.config_dir / self.config_box.get())
        raw_input = self.input_var.get().strip().strip('"')
        mock = self.mock_var.get()
        return GenerationRequest(
            config_path=config_path,
            input_path=Path(raw_input) if raw_input else None,
            mock=mock,
            notes=self.notes_text(),
            output_dir=self.output_dir,
            history_db=self.history_db if self.history_var.get() and not mock else None,
        )

    @property
    def is_running(self) -> bool:
        return self._worker is not None and self._worker.is_alive()

    def start_generation(self) -> None:
        if self.is_running:
            return
        if not self.config_box.get():
            messagebox.showwarning("Нет конфига", "Выберите конфиг клиента.", parent=self)
            return
        request = self.build_request()
        problem = validate_request(request)
        if problem:
            messagebox.showwarning("Проверьте данные", problem, parent=self)
            return

        self._set_running(True)
        self._last_report = None
        self.open_report_button.configure(state="disabled")
        self.progress.set(0)
        self._set_status("Запуск сборки…")
        self._append_log(logging.INFO, "— Сборка отчёта —")
        self._worker = threading.Thread(
            target=run_generation, args=(request, self._events), name="report-generation", daemon=True
        )
        self._worker.start()

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        self.generate_button.configure(
            state=state, text="Идёт сборка…" if running else "Сгенерировать презентацию (.pptx)"
        )
        self.config_box.configure(state="disabled" if running else "readonly")
        self.mock_checkbox.configure(state=state)
        if running:
            for widget in (self.input_entry, self.browse_button, self.history_checkbox):
                widget.configure(state="disabled")
        else:
            self._sync_source_state()

    # --- События рабочего потока ---------------------------------------------------

    def _poll_events(self) -> None:
        try:
            for _ in range(_MAX_EVENTS_PER_POLL):
                self._handle_event(self._events.get_nowait())
        except queue.Empty:
            pass
        finally:
            self._poll_job = self.after(_POLL_MS, self._poll_events)

    def _handle_event(self, event: Event) -> None:
        match event:
            case ProgressEvent():
                self.progress.set(event.fraction)
                self._set_status(f"Этап {event.step} из {event.total}: {event.label}…")
            case LogEvent():
                self._append_log(event.level, event.message)
            case SuccessEvent():
                self._on_success(event.result)
            case FailureEvent():
                self._on_failure(event)

    def _on_success(self, result: ReportResult) -> None:
        self._set_running(False)
        self._last_report = result.output_path
        self.progress.set(1)
        self._set_status(f"Готово: {result.output_path.name} ({result.timing_summary()})")
        self.open_report_button.configure(state="normal")
        SuccessDialog(self, result)

    def _on_failure(self, event: FailureEvent) -> None:
        self._set_running(False)
        self.progress.set(0)
        self._set_status(f"Ошибка: {event.title}")
        messagebox.showerror(event.title, event.message, parent=self)

    def _set_status(self, text: str) -> None:
        self.status_label.configure(text=text)

    def _append_log(self, level: int, message: str) -> None:
        self.log_box.configure(state="normal")
        prefix = "⚠ " if level >= logging.WARNING else ""
        self.log_box.insert("end", f"{prefix}{message}\n")
        lines = int(self.log_box.index("end-1c").split(".")[0])
        if lines > _MAX_LOG_LINES:
            self.log_box.delete("1.0", f"{lines - _MAX_LOG_LINES}.0")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    # --- Открытие результатов -------------------------------------------------------

    def open_report(self) -> None:
        if self._last_report is not None:
            self._open(self._last_report)

    def open_folder(self) -> None:
        folder = self._last_report.parent if self._last_report is not None else self.output_dir
        folder.mkdir(parents=True, exist_ok=True)
        self._open(folder)

    def _open(self, path: Path) -> None:
        try:
            open_path(path)
        except OSError as exc:
            messagebox.showerror("Не удалось открыть", f"{path}\n\n{exc}", parent=self)

    # --- Оформление и закрытие ------------------------------------------------------

    def _change_appearance(self, label: str) -> None:
        ctk.set_appearance_mode(_APPEARANCE_MODES[label])

    def report_callback_exception(self, exc_type, exc_value, exc_traceback) -> None:  # noqa: ANN001 — сигнатура tkinter
        """Ошибка в обработчике кнопки: в собранном .exe нет консоли, поэтому показываем её окном."""
        logger.error("Ошибка интерфейса", exc_info=(exc_type, exc_value, exc_traceback))
        messagebox.showerror("Ошибка интерфейса", f"{exc_type.__name__}: {exc_value}", parent=self)

    def _on_close(self) -> None:
        if self.is_running and not messagebox.askyesno(
            "Отчёт ещё собирается", "Закрыть приложение? Незавершённый отчёт не сохранится.", parent=self
        ):
            return
        self.after_cancel(self._poll_job)
        self._source_logger.removeHandler(self._log_handler)
        self.destroy()


class SuccessDialog(ctk.CTkToplevel):
    """Уведомление о готовом отчёте с кнопками «Открыть отчёт» и «Открыть папку»."""

    def __init__(self, app: ReportApp, result: ReportResult) -> None:
        super().__init__(app)
        self.app = app
        self.title("Отчёт готов")
        self.resizable(False, False)
        self.transient(app)

        date_from, date_to = result.period
        lines = [
            f"Период: {date_from:%d.%m.%Y} – {date_to:%d.%m.%Y}",
            f"Кампаний: {result.campaigns_count} · слайдов: {result.slides_count}",
        ]
        if result.report_month is not None:
            mom = f"динамика к {result.compared_to}" if result.compared_to else "прошлого месяца в истории нет"
            saved = "сохранены в историю" if result.history_saved else "не сохранены в историю (см. журнал)"
            lines.append(f"Итоги {result.report_month} {saved}; {mom}")

        ctk.CTkLabel(self, text="Презентация сохранена", font=ctk.CTkFont(size=18, weight="bold"), anchor="w").pack(
            fill="x", padx=20, pady=(18, 4)
        )
        ctk.CTkLabel(self, text=str(result.output_path), anchor="w", justify="left", wraplength=520).pack(
            fill="x", padx=20
        )
        ctk.CTkLabel(self, text="\n".join(lines), anchor="w", justify="left", text_color=("gray30", "gray70")).pack(
            fill="x", padx=20, pady=(8, 12)
        )
        buttons = ctk.CTkFrame(self, fg_color="transparent")
        buttons.pack(fill="x", padx=20, pady=(0, 18))
        ctk.CTkButton(buttons, text="Открыть отчёт", command=self._open_report).pack(side="left")
        ctk.CTkButton(buttons, text="Открыть папку", command=self._open_folder).pack(side="left", padx=8)
        ctk.CTkButton(
            buttons, text="Закрыть", width=90, fg_color="transparent", border_width=1,
            text_color=("gray10", "gray90"), command=self.destroy,
        ).pack(side="right")  # fmt: skip

        self.bind("<Escape>", lambda _event: self.destroy())
        self.after(50, self._show_on_top)

    def _show_on_top(self) -> None:
        """Над главным окном и модально: CTkToplevel в Windows иначе может открыться позади."""
        self.update_idletasks()
        x = self.app.winfo_rootx() + (self.app.winfo_width() - self.winfo_width()) // 2
        y = self.app.winfo_rooty() + (self.app.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.lift()
        self.focus_force()
        try:
            self.grab_set()
        except tk.TclError:  # окно ещё не показано оконным менеджером — остаётся немодальным
            logger.debug("Не удалось сделать окно «Отчёт готов» модальным")

    def _open_report(self) -> None:
        self.app.open_report()
        self.destroy()

    def _open_folder(self) -> None:
        self.app.open_folder()
        self.destroy()


def run() -> int:
    """Запуск окна: первый старт .exe раскладывает конфиги из бандла в config/."""
    notices = []
    try:
        copied = ensure_user_configs()
    except OSError as exc:
        notices.append(f"Не удалось скопировать конфиги по умолчанию в {CONFIG_DIR}: {exc}")
    else:
        if copied:
            notices.append(f"Конфиги по умолчанию скопированы в {CONFIG_DIR}: {', '.join(p.name for p in copied)}")
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    app = ReportApp(notices=notices)
    app.mainloop()
    return 0
