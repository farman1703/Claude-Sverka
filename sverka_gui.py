#!/usr/bin/env python3
"""
Графическое приложение для сверки двух Excel-файлов.

Вкладки окна:
    Новая сверка             выбор системы, отчёта, периода и двух файлов
    Результат                итог, разбивка по месяцам, расхождения, Excel-отчёт
    История                  все прошлые сверки с отбором
    Справочники              системы и их отчёты
    Дополнительные параметры те же настройки, что у compare_excel.py в командной строке

Сама сверка и построение отчёта целиком выполняются модулем compare_excel.py.
Библиотеку pandas, которую он использует, в приложении заменяет pandas_lite
(чистый Python): скомпилированные файлы pandas и numpy Windows может
заблокировать. Сам compare_excel.py при этом не меняется.

Запуск:
    python sverka_gui.py
"""

import json
import queue
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import service
    from engine import ce
    from gui_dicts import DictView
    from gui_history import HistoryView
    from gui_result import ResultView
    from gui_widgets import BG, CARD, MUTED, fix_tree_colors, install_wheel, make_fonts
    from storage import Store
    IMPORT_ERROR = None
except ImportError as error:        # не установлены библиотеки: сообщим об этом в окне
    IMPORT_ERROR = error

APP_TITLE = "Сверка Excel-файлов"
EXCEL_TYPES = [("Файлы Excel", "*.xlsx *.xlsm"), ("Все файлы", "*.*")]
NO_SYSTEM = "— без справочника —"


def split_list(text):
    """«а; б; в» -> ["а", "б", "в"]. Названия столбцов и листов могут содержать пробелы."""
    return [part.strip() for part in text.split(";") if part.strip()]


def parse_float(text, title):
    try:
        return float(text.strip().replace(",", "."))
    except ValueError:
        raise ValueError(f"{title}: нужно число, а указано «{text.strip()}»") from None


class App(tk.Tk):
    def __init__(self, store=None):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("980x700")
        self.minsize(900, 600)
        self.store = store or Store()
        self.fonts = make_fonts(self)
        self.events = queue.Queue()
        self.last_dir = str(Path.home())
        self.system_ids, self.report_ids = {}, {}

        style = ttk.Style(self)
        if self.tk.call("tk", "windowingsystem") == "x11" and "clam" in style.theme_names():
            style.theme_use("clam")
        fix_tree_colors(style)
        self.configure(bg=BG)
        for name in ("TFrame", "TLabel", "TCheckbutton", "TRadiobutton", "TLabelframe",
                     "TLabelframe.Label"):
            style.configure(name, background=BG)
        style.configure("Run.TButton", font=self.fonts["head"], padding=(18, 7))
        style.configure("Hint.TLabel", foreground=MUTED)
        style.configure("Link.TButton", relief="flat", padding=(2, 1), foreground="#1F5FA8",
                        background=CARD)
        style.configure("Treeview", rowheight=24)
        install_wheel(self)

        self._build()
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.after(100, self._poll)

    # ---------- окно ----------

    def _build(self):
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=8, pady=8)
        self.new_tab = ttk.Frame(self.tabs, padding=16)
        self.result = ResultView(self.tabs, self)
        self.history = HistoryView(self.tabs, self)
        self.dicts = DictView(self.tabs, self)
        options = ttk.Frame(self.tabs, padding=16)
        for tab, title in ((self.new_tab, "Новая сверка"), (self.result, "Результат"),
                           (self.history, "История"), (self.dicts, "Справочники"),
                           (options, "Дополнительные параметры")):
            self.tabs.add(tab, text=f"  {title}  ")
        self._build_new(self.new_tab)
        self._build_options(options)
        self.options_tab = options

    def _build_new(self, main):
        main.columnconfigure(1, weight=1)
        ttk.Label(main, style="Hint.TLabel", wraplength=900, justify="left", text=(
            "Выберите два Excel-файла с одинаковой структурой. Первый файл считается "
            "эталоном: результат показывает, что с его данными произошло во втором файле."
        )).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        # --- что сверяем ---
        self.system, self.report = tk.StringVar(value=NO_SYSTEM), tk.StringVar()
        ttk.Label(main, text="Система:").grid(row=1, column=0, sticky="w", pady=4)
        what = ttk.Frame(main)
        what.grid(row=1, column=1, columnspan=2, sticky="ew", padx=(8, 0), pady=4)
        self.system_box = ttk.Combobox(what, textvariable=self.system, state="readonly", width=28)
        self.system_box.pack(side="left")
        self.system_box.bind("<<ComboboxSelected>>", lambda e: self.fill_reports())
        ttk.Label(what, text="Отчёт:").pack(side="left", padx=(18, 8))
        self.report_box = ttk.Combobox(what, textvariable=self.report, state="readonly", width=40)
        self.report_box.pack(side="left", fill="x", expand=True)

        self.date_from, self.date_to = tk.StringVar(), tk.StringVar()
        ttk.Label(main, text="Период:").grid(row=2, column=0, sticky="w", pady=4)
        period = ttk.Frame(main)
        period.grid(row=2, column=1, columnspan=2, sticky="ew", padx=(8, 0), pady=4)
        ttk.Label(period, text="с").pack(side="left")
        ttk.Entry(period, textvariable=self.date_from, width=12).pack(side="left", padx=(6, 10))
        ttk.Label(period, text="по").pack(side="left")
        ttk.Entry(period, textvariable=self.date_to, width=12).pack(side="left", padx=(6, 16))
        for kind, text in (("month", "Прошлый месяц"), ("quarter", "Прошлый квартал"),
                           ("year", "С начала года"), ("clear", "Очистить")):
            ttk.Button(period, text=text, command=lambda k=kind: self.set_period(k)).pack(
                side="left", padx=(0, 6))
        ttk.Label(main, style="Hint.TLabel", wraplength=900, justify="left", text=(
            "Систему, отчёт и период указывать не обязательно. Даты пишутся так: 31.12.2026. "
            "Если период не указать, он берётся из дат в самих файлах.")).grid(
                row=3, column=0, columnspan=3, sticky="w", pady=(2, 14))

        # --- файлы ---
        self.file1, self.file2 = tk.StringVar(), tk.StringVar()
        for row, (label, var) in enumerate((("Файл 1 (эталон):", self.file1),
                                            ("Файл 2:", self.file2)), 4):
            ttk.Label(main, text=label).grid(row=row, column=0, sticky="w", pady=4)
            entry = ttk.Entry(main, textvariable=var)
            entry.grid(row=row, column=1, sticky="ew", padx=8, pady=4)
            var.trace_add("write", lambda *_, e=entry: e.after_idle(e.xview_moveto, 1))
            ttk.Button(main, text="Выбрать…", command=lambda v=var: self._choose(v)).grid(
                row=row, column=2, pady=4)

        # --- запуск ---
        run = ttk.Frame(main)
        run.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(16, 8))
        run.columnconfigure(1, weight=1)
        self.run_button = ttk.Button(run, text="Сверить", style="Run.TButton", command=self.start)
        self.run_button.grid(row=0, column=0)
        self.progress = ttk.Progressbar(run, mode="indeterminate")
        self.status = ttk.Label(main, text="", style="Hint.TLabel", wraplength=900, justify="left")
        self.status.grid(row=7, column=0, columnspan=3, sticky="w")
        ttk.Label(main, style="Hint.TLabel", wraplength=900, justify="left", text=(
            "После сверки откроется вкладка «Результат». Каждая сверка сохраняется на вкладке "
            "«История»: её можно открыть позже и заново сохранить Excel-отчёт. Ваши исходные "
            "файлы программа не меняет и никуда не копирует.")).grid(
                row=8, column=0, columnspan=3, sticky="w", pady=(18, 0))
        self.fill_systems()

    def _build_options(self, box):
        """Те же параметры, что у compare_excel.py в командной строке."""
        box.columnconfigure(1, weight=1)
        ttk.Label(box, style="Hint.TLabel", wraplength=900, justify="left", text=(
            "Менять эти параметры не обязательно: пока они не тронуты, сверка идёт "
            "с настройками по умолчанию.")).grid(row=0, column=0, columnspan=3,
                                                sticky="w", pady=(0, 10))
        self.name_a, self.name_b = tk.StringVar(), tk.StringVar()
        self.keys, self.ignore = tk.StringVar(), tk.StringVar()
        self.sheets, self.sums = tk.StringVar(), tk.StringVar()
        self.header, self.tolerance, self.similarity = (tk.StringVar() for _ in range(3))
        self.skip_first = tk.BooleanVar()
        self.by_order = tk.BooleanVar()
        self.ignore_case = tk.BooleanVar()
        self.text_numbers = tk.BooleanVar()
        self.full = tk.BooleanVar()
        self.reset_options()

        def line(row, label, hint):
            ttk.Label(box, text=label).grid(row=row, column=0, sticky="w", pady=4)
            ttk.Label(box, text=hint, style="Hint.TLabel").grid(row=row, column=2,
                                                                sticky="w", pady=4)

        line(1, "Названия файлов в отчёте:", "по умолчанию берутся из имён файлов")
        names = ttk.Frame(box)
        names.grid(row=1, column=1, sticky="ew", padx=8)
        names.columnconfigure((0, 1), weight=1)
        ttk.Entry(names, textvariable=self.name_a).grid(row=0, column=0, sticky="ew")
        ttk.Entry(names, textvariable=self.name_b).grid(row=0, column=1, sticky="ew",
                                                        padx=(6, 0))
        rows = ((self.keys, "Ключевые столбцы:", "названия столбцов через «;»"),
                (self.ignore, "Не сравнивать столбцы:", "названия столбцов через «;»"),
                (self.sheets, "Сравнить только листы:", "названия листов через «;»"),
                (self.sums, "Суммы по столбцам:", "буквы Excel: AE AH AK:AM или Лист!C:E"))
        for row, (var, label, hint) in enumerate(rows, 2):
            line(row, label, hint)
            ttk.Entry(box, textvariable=var).grid(row=row, column=1, sticky="ew", padx=8)

        numbers = ttk.Frame(box)
        numbers.grid(row=6, column=0, columnspan=3, sticky="w", pady=(10, 8))
        ttk.Label(numbers, text="Строка заголовков:").pack(side="left")
        ttk.Spinbox(numbers, textvariable=self.header, from_=1, to=1000, width=5).pack(
            side="left", padx=(6, 18))
        ttk.Label(numbers, text="Допуск для чисел:").pack(side="left")
        ttk.Entry(numbers, textvariable=self.tolerance, width=9).pack(side="left",
                                                                     padx=(6, 18))
        ttk.Label(numbers, text="Доля сходства строк (0–1):").pack(side="left")
        ttk.Entry(numbers, textvariable=self.similarity, width=6).pack(side="left", padx=6)

        checks = ttk.Frame(box)
        checks.grid(row=7, column=0, columnspan=3, sticky="w")
        flags = ((self.skip_first, "Не сравнивать столбец A (нумерация строк)"),
                 (self.by_order, "Сравнивать строки строго по порядку"),
                 (self.ignore_case, "Не учитывать регистр в тексте"),
                 (self.text_numbers, "Текст, похожий на число, считать числом"),
                 (self.full, "Добавить в Excel-отчёт подробные листы"))
        for k, (var, label) in enumerate(flags):
            ttk.Checkbutton(checks, text=label, variable=var).grid(
                row=k, column=0, sticky="w", pady=3)
        ttk.Button(box, text="Сбросить параметры", command=self.reset_options).grid(
            row=8, column=0, columnspan=3, sticky="w", pady=(14, 0))

    def reset_options(self):
        """Значения по умолчанию - как у скрипта без параметров."""
        for var in (self.name_a, self.name_b, self.keys, self.ignore, self.sheets, self.sums):
            var.set("")
        self.header.set("1")
        self.tolerance.set("0")
        self.similarity.set("0.5")
        for var in (self.skip_first, self.by_order, self.ignore_case, self.text_numbers,
                    self.full):
            var.set(False)

    # ---------- система, отчёт, период ----------

    def fill_systems(self):
        """Список систем из справочника; первая строка - сверка без справочника."""
        self.system_ids = {s["name"]: s["id"] for s in self.store.systems()}
        self.system_box.configure(values=[NO_SYSTEM] + list(self.system_ids))
        if self.system.get() not in self.system_ids:
            self.system.set(NO_SYSTEM)
        self.fill_reports()

    def fill_reports(self):
        """В списке отчётов - только отчёты выбранной системы."""
        system_id = self.system_ids.get(self.system.get())
        rows = self.store.reports(system_id) if system_id else []
        self.report_ids = {r["name"]: r["id"] for r in rows}
        self.report_box.configure(values=list(self.report_ids),
                                  state="readonly" if rows else "disabled")
        if self.report.get() not in self.report_ids:
            self.report.set("")

    def dictionaries_changed(self):
        self.fill_systems()
        self.history.refresh()

    def set_period(self, kind):
        if kind == "clear":
            self.date_from.set("")
            self.date_to.set("")
            return
        first, last = service.quick_period(kind)
        self.date_from.set(first.strftime("%d.%m.%Y"))
        self.date_to.set(last.strftime("%d.%m.%Y"))

    def read_period(self):
        """Период из полей окна: (с, по) или (None, None), если он не указан."""
        first = service.parse_user_date(self.date_from.get())
        last = service.parse_user_date(self.date_to.get())
        if (first is None) != (last is None):
            raise ValueError("Укажите обе даты периода или оставьте обе пустыми.")
        if first and first > last:
            raise ValueError("Дата «с» позже даты «по».")
        return first, last

    def read_report(self):
        """Выбранный отчёт справочника или None (сверка без справочника)."""
        if self.system.get() not in self.system_ids:
            return None
        if not self.report_ids:
            raise ValueError("У этой системы пока нет отчётов. Добавьте их на вкладке "
                             "«Справочники» или выберите «без справочника».")
        if self.report.get() not in self.report_ids:
            raise ValueError("Выберите отчёт.")
        return self.store.report(self.report_ids[self.report.get()])

    def _choose(self, var):
        current = var.get().strip().strip('"')
        folder = str(Path(current).parent) if current else self.last_dir
        path = filedialog.askopenfilename(parent=self, title="Выберите Excel-файл",
                                          initialdir=folder, filetypes=EXCEL_TYPES)
        if path:
            var.set(path)
            self.last_dir = str(Path(path).parent)

    def _set_busy(self, busy):
        self.run_button.configure(state="disabled" if busy else "normal")
        if busy:
            self.progress.grid(row=0, column=1, sticky="ew", padx=(14, 0))
            self.progress.start(12)
        else:
            self.progress.stop()
            self.progress.grid_remove()

    # ---------- параметры ----------

    def read_options(self):
        """Собирает Options так же, как это делает main() в compare_excel.py."""
        try:
            header = int(self.header.get().strip())
        except ValueError:
            raise ValueError("Строка заголовков: нужен номер строки, например 1") from None
        if header < 1:
            raise ValueError("Строка заголовков: номер строки начинается с 1")
        names = (self.name_a.get().strip(), self.name_b.get().strip())
        if any(names) and not all(names):
            raise ValueError("Названия файлов в отчёте: заполните оба поля или оставьте "
                             "оба пустыми")
        o = ce.Options(
            keys=split_list(self.keys.get()) or None,
            by_order=self.by_order.get(),
            header=header - 1,
            tolerance=parse_float(self.tolerance.get(), "Допуск для чисел"),
            ignore_case=self.ignore_case.get(),
            text_numbers=self.text_numbers.get(),
            similarity=parse_float(self.similarity.get(), "Доля сходства строк"),
            ignore=tuple(split_list(self.ignore.get())),
            skip_first_column=self.skip_first.get(),
            sheets=split_list(self.sheets.get()) or None,
            full=self.full.get())
        o.sum_cols = ce.parse_columns(split_list(self.sums.get()))
        return o, (list(names) if all(names) else None)

    # ---------- сверка ----------

    def start(self):
        paths = []
        for var, title in ((self.file1, "Файл 1"), (self.file2, "Файл 2")):
            text = var.get().strip().strip('"')
            if not text:
                messagebox.showwarning(APP_TITLE, f"{title} не выбран", parent=self)
                return
            if not Path(text).is_file():
                messagebox.showerror(APP_TITLE, f"Файл не найден: {text}", parent=self)
                return
            paths.append(Path(text))
        try:
            report = self.read_report()
            period = self.read_period()
        except ValueError as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
            return
        try:
            o, names = self.read_options()
        except ValueError as e:
            self.tabs.select(self.options_tab)
            messagebox.showerror(APP_TITLE, str(e), parent=self)
            return

        saved = json.dumps({"keys": o.keys, "ignore": list(o.ignore), "sheets": o.sheets,
                            "header": o.header + 1, "tolerance": o.tolerance,
                            "similarity": o.similarity, "by_order": o.by_order,
                            "skip_first_column": o.skip_first_column, "full": o.full,
                            "ignore_case": o.ignore_case, "text_numbers": o.text_numbers,
                            "sum": self.sums.get().strip(), "names": names}, ensure_ascii=False)
        run_id = self.store.create_run(report, *paths, *period, saved)
        self.status.configure(text="Идёт сверка…")
        self._set_busy(True)
        threading.Thread(target=self._work, daemon=True, args=(
            run_id, *paths, report["id"] if report else None, *period, o, names)).start()

    def _work(self, run_id, file1, file2, report_id, first, last, o, names):
        """Фоновый поток: сверка с записью результата в историю."""
        try:
            store = Store(self.store.folder)       # у потока своё подключение к базе
            try:
                report = store.report(report_id) if report_id else None
                service.execute(store, run_id, file1, file2, report, first, last, o, names)
            finally:
                store.close()
        except Exception as e:                     # сбой вне самой сверки
            self.events.put(("crash", run_id, f"({type(e).__name__}): {e}"))
            return
        self.events.put(("done", run_id, None))

    def _poll(self):
        try:
            while True:
                kind, run_id, text = self.events.get_nowait()
                self._set_busy(False)
                self.status.configure(text="")
                if kind == "crash":
                    messagebox.showerror(APP_TITLE, "Не удалось записать сверку: " + text,
                                         parent=self)
                self.history.refresh()
                self.dicts.refresh()
                self.fill_reports()
                self.show_run(run_id)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    # ---------- переходы между вкладками ----------

    def show_run(self, run_id):
        """Открывает сверку на вкладке «Результат»."""
        self.result.show(run_id)
        self.tabs.select(self.result)

    def prefill(self, run):
        """«Сверить заново»: система, отчёт, период и файлы подставляются из сверки."""
        self.fill_systems()
        self.system.set(run["system"] if run["system"] in self.system_ids else NO_SYSTEM)
        self.fill_reports()
        if run["report"] in self.report_ids:
            self.report.set(run["report"])
        if run["period_from"] and not run["period_auto"]:
            self.date_from.set(service.format_day(run["period_from"]))
            self.date_to.set(service.format_day(run["period_to"]))
        else:
            self.set_period("clear")
        self.file1.set(run["file_a"] or "")
        self.file2.set(run["file_b"] or "")
        self.tabs.select(self.new_tab)

    def close(self):
        self.store.close()
        self.destroy()

    def report_callback_exception(self, exc, val, tb):
        """Ошибка в самом окне: показываем, а не теряем (консоли может не быть)."""
        text = "".join(traceback.format_exception(exc, val, tb))
        messagebox.showerror(APP_TITLE, "Непредвиденная ошибка:\n\n" + text[-1500:],
                             parent=self)


def main():
    if sys.platform == "win32":   # чёткий текст на экранах с масштабированием
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    problem = None
    if IMPORT_ERROR is not None:
        problem = (f"Не удалось загрузить программу: {IMPORT_ERROR}\n\n"
                   "Запустите установку (install) ещё раз.")
    else:
        try:
            store = Store()
        except Exception as e:
            problem = f"Не удалось открыть папку с данными программы: {e}"
    if problem:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_TITLE, problem)
        return 2
    App(store).mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
