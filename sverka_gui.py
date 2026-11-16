#!/usr/bin/env python3
"""
Графическое приложение для сверки двух Excel-файлов.

Это только окно: выбор файлов, запуск сверки и сохранение отчёта. Сама сверка
и построение отчёта целиком выполняются модулем compare_excel.py - приложение
вызывает ту же функцию run_pair, что и запуск из командной строки, с теми же
значениями параметров по умолчанию.

Библиотеку pandas, которую использует compare_excel.py, в приложении заменяет
модуль pandas_lite (чистый Python): скомпилированные файлы pandas и numpy
Windows может заблокировать. Сам compare_excel.py при этом не меняется.

Запуск:
    python sverka_gui.py
"""

import atexit
import contextlib
import io
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "Сверка Excel-файлов"
EXCEL_TYPES = [("Файлы Excel", "*.xlsx *.xlsm"), ("Все файлы", "*.*")]
GREEN, GREEN_TEXT = "#C6EFCE", "#006100"
RED, RED_TEXT = "#FFC7CE", "#9C0006"
GREY_TEXT = "#595959"


def split_list(text):
    """«а; б; в» -> ["а", "б", "в"]. Названия столбцов и листов могут содержать пробелы."""
    return [part.strip() for part in text.split(";") if part.strip()]


def parse_float(text, title):
    try:
        return float(text.strip().replace(",", "."))
    except ValueError:
        raise ValueError(f"{title}: нужно число, а указано «{text.strip()}»") from None


def open_file(path):
    """Открывает файл программой по умолчанию (Excel)."""
    if sys.platform == "win32":
        os.startfile(path)   # noqa: S606 - открываем отчёт, который сами создали
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class App(tk.Tk):
    def __init__(self, engine):
        super().__init__()
        self.engine = engine
        self.title(APP_TITLE)
        self.geometry("900x640")
        self.minsize(760, 520)

        # Отчёты складываются во временную папку; пользователь сохраняет копию
        self.tmp = Path(tempfile.mkdtemp(prefix="sverka_"))
        atexit.register(shutil.rmtree, self.tmp, ignore_errors=True)
        self.events = queue.Queue()
        self.report = None       # готовый отчёт во временной папке
        self.saved = None        # куда пользователь его сохранил
        self.last_dir = str(Path.home())

        style = ttk.Style(self)
        if self.tk.call("tk", "windowingsystem") == "x11" and "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Run.TButton", font=("TkDefaultFont", 11, "bold"), padding=(18, 7))
        style.configure("Hint.TLabel", foreground=GREY_TEXT)

        self._build()
        self.after(100, self._poll)

    # ---------- окно ----------

    def _build(self):
        # Две вкладки: сама сверка и необязательные параметры
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=8, pady=8)
        main = ttk.Frame(self.tabs, padding=14)
        main.columnconfigure(1, weight=1)
        self.tabs.add(main, text="  Сверка  ")
        options = ttk.Frame(self.tabs, padding=14)
        self.tabs.add(options, text="  Дополнительные параметры  ")
        self._build_options(options)

        ttk.Label(main, style="Hint.TLabel", text=(
            "Выберите два Excel-файла с одинаковой структурой. Первый файл считается "
            "эталоном: отчёт показывает, что с его данными произошло во втором файле."),
            wraplength=820, justify="left").grid(row=0, column=0, columnspan=3,
                                                 sticky="w", pady=(0, 10))

        self.file1, self.file2 = tk.StringVar(), tk.StringVar()
        for row, (label, var) in enumerate((("Файл 1 (эталон):", self.file1),
                                            ("Файл 2:", self.file2)), 1):
            ttk.Label(main, text=label).grid(row=row, column=0, sticky="w", pady=3)
            entry = ttk.Entry(main, textvariable=var)
            entry.grid(row=row, column=1, sticky="ew", padx=8, pady=3)
            var.trace_add("write", lambda *_, e=entry: e.after_idle(e.xview_moveto, 1))
            ttk.Button(main, text="Выбрать…", command=lambda v=var: self._choose(v)).grid(
                row=row, column=2, pady=3)

        # --- запуск ---
        run = ttk.Frame(main)
        run.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(12, 8))
        run.columnconfigure(1, weight=1)
        self.run_button = ttk.Button(run, text="Сверить", style="Run.TButton",
                                     command=self.start)
        self.run_button.grid(row=0, column=0)
        self.progress = ttk.Progressbar(run, mode="indeterminate")

        # --- результат ---
        self.verdict = tk.Label(main, text="", anchor="w", padx=10, pady=7,
                                font=("TkDefaultFont", 12, "bold"))
        self.verdict.grid(row=6, column=0, columnspan=3, sticky="ew")
        self.summary = ttk.Label(main, text="", style="Hint.TLabel")
        self.summary.grid(row=7, column=0, columnspan=3, sticky="w", pady=(3, 6))

        log = ttk.Frame(main)
        log.grid(row=8, column=0, columnspan=3, sticky="nsew")
        main.rowconfigure(8, weight=1, minsize=90)
        log.columnconfigure(0, weight=1)
        log.rowconfigure(0, weight=1)
        self.log = tk.Text(log, height=10, wrap="none", state="disabled",
                           font="TkFixedFont", relief="solid", borderwidth=1)
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll_y = ttk.Scrollbar(log, orient="vertical", command=self.log.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(log, orient="horizontal", command=self.log.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        self.log.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        bottom = ttk.Frame(main)
        bottom.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        self.save_button = ttk.Button(bottom, text="Сохранить отчёт…", command=self.save,
                                      state="disabled")
        self.save_button.grid(row=0, column=0)
        self.open_button = ttk.Button(bottom, text="Открыть отчёт", command=self.open_report,
                                      state="disabled")
        self.open_button.grid(row=0, column=1, padx=8)
        self.status = ttk.Label(main, text="", style="Hint.TLabel", wraplength=840,
                                justify="left")
        self.status.grid(row=10, column=0, columnspan=3, sticky="w", pady=(6, 0))

    def _build_options(self, box):
        """Те же параметры, что у compare_excel.py в командной строке."""
        box.columnconfigure(1, weight=1)
        ttk.Label(box, style="Hint.TLabel", wraplength=820, justify="left", text=(
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
                 (self.full, "Добавить в отчёт подробные листы"))
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

    def _choose(self, var):
        current = var.get().strip().strip('"')
        folder = str(Path(current).parent) if current else self.last_dir
        path = filedialog.askopenfilename(parent=self, title="Выберите Excel-файл",
                                          initialdir=folder, filetypes=EXCEL_TYPES)
        if path:
            var.set(path)
            self.last_dir = str(Path(path).parent)

    def _set_log(self, text):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.insert("1.0", text)
        self.log.configure(state="disabled")

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
        eng = self.engine
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
        o = eng.Options(
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
        o.sum_cols = eng.parse_columns(split_list(self.sums.get()))
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
            o, names = self.read_options()
        except ValueError as e:
            self.tabs.select(1)
            messagebox.showerror(APP_TITLE, str(e), parent=self)
            return

        self.report = self.saved = None
        self.save_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.verdict.configure(text="", bg=self.cget("bg"))
        self.summary.configure(text="")
        self.status.configure(text="Идёт сверка…")
        self._set_log("")
        self._set_busy(True)
        out_dir = Path(tempfile.mkdtemp(dir=self.tmp))
        threading.Thread(target=self._work, args=(*paths, o, names, out_dir),
                         daemon=True).start()

    def _work(self, file1, file2, o, names, out_dir):
        """Фоновый поток: вызывает run_pair из compare_excel.py."""
        eng = self.engine
        output = io.StringIO()
        try:
            # Имя отчёта - как у запуска из командной строки без параметра -o
            nm = eng.derive_names(file1, file2, names)
            report = out_dir / (f"{nm.prefix}_differences.xlsx" if nm.prefix
                                else "differences.xlsx")
            with contextlib.redirect_stdout(output):
                res, nm, report = eng.run_pair(file1, file2, o, names, report)
            self.events.put(("done", res, report, output.getvalue()))
        except PermissionError as e:
            self.events.put(("error", f"Нет доступа к файлу {e.filename}. Если он открыт "
                                      "в Excel, закройте его и запустите сверку снова"))
        except (ValueError, OSError) as e:
            self.events.put(("error", str(e)))
        except Exception as e:   # непредвиденная ошибка
            self.events.put(("error", f"({type(e).__name__}): {e}"))

    def _poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                self._set_busy(False)
                if event[0] == "done":
                    self._show_result(*event[1:])
                else:
                    self.status.configure(text="")
                    self.verdict.configure(text="  ОШИБКА", bg=RED, fg=RED_TEXT)
                    self._set_log("Ошибка: " + event[1])
                    messagebox.showerror(APP_TITLE, "Ошибка: " + event[1], parent=self)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _show_result(self, res, report, output):
        self.report = report
        t = res.totals
        if res.has_differences:
            self.verdict.configure(text="  ИТОГ: ЕСТЬ РАЗЛИЧИЯ", bg=RED, fg=RED_TEXT)
        else:
            self.verdict.configure(text="  ИТОГ: ФАЙЛЫ СОВПАДАЮТ", bg=GREEN, fg=GREEN_TEXT)
        self.summary.configure(text=(
            f"Строк в файлах: {t['rows_a']} / {t['rows_b']}   ·   совпали полностью: "
            f"{t['exact']}   ·   с расхождениями: {t['changed']}   ·   только в одном "
            f"из файлов: {t['only_a'] + t['only_b']}"))
        # Тот же текст, что скрипт печатает в консоль (без пути во временной папке
        # и без совета поставить python-calamine: приложению она не нужна)
        lines = [s for s in output.strip("\n").split("\n")
                 if not s.startswith(("Отчёт: ", "  Чтение файлов можно ускорить"))]
        self._set_log("\n".join(lines))
        self.save_button.configure(state="normal")
        self.open_button.configure(state="normal")
        self.status.configure(text=f"Отчёт готов: {report.name}. Сохраните его кнопкой слева.")

    # ---------- отчёт ----------

    def save(self):
        if not self.report:
            return
        path = filedialog.asksaveasfilename(
            parent=self, title="Сохранить отчёт", initialdir=self.last_dir,
            initialfile=self.report.name, defaultextension=".xlsx",
            filetypes=[("Файлы Excel", "*.xlsx")])
        if not path:
            return
        try:
            shutil.copyfile(self.report, path)
        except PermissionError:
            messagebox.showerror(APP_TITLE, f"Нет доступа к файлу {path}. Если он открыт "
                                 "в Excel, закройте его и сохраните отчёт снова", parent=self)
            return
        except OSError as e:
            messagebox.showerror(APP_TITLE, f"Не удалось сохранить отчёт: {e}", parent=self)
            return
        self.saved = Path(path)
        self.last_dir = str(self.saved.parent)
        self.status.configure(text=f"Отчёт сохранён: {self.saved}")

    def open_report(self):
        target = self.saved or self.report
        if not target:
            return
        try:
            open_file(target)
        except OSError as e:
            messagebox.showerror(APP_TITLE, f"Не удалось открыть отчёт: {e}", parent=self)

    def report_callback_exception(self, exc, val, tb):
        """Ошибка в самом окне: показываем, а не теряем (консоли может не быть)."""
        text = "".join(traceback.format_exception(exc, val, tb))
        messagebox.showerror(APP_TITLE, "Непредвиденная ошибка:\n\n" + text[-1500:],
                             parent=self)


def load_engine():
    """Модуль сверки compare_excel.py; в роли pandas ему подставляется pandas_lite."""
    import pandas_lite
    sys.modules["pandas"] = pandas_lite
    import compare_excel
    return compare_excel


def main():
    if sys.platform == "win32":   # чёткий текст на экранах с масштабированием
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    try:
        engine = load_engine()
    except ImportError as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_TITLE, f"Не удалось загрузить модуль сверки: {e}\n\n"
                             "Запустите установку (install) ещё раз.")
        return 2
    App(engine).mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
