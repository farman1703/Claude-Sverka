"""Вкладка «Результат»: итог сверки, разбивка по месяцам, расхождения, Excel-отчёт."""

import json
import shutil
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import service
from gui_widgets import (BAD_BG, BAD_FG, BG, CARD, GREY_BG, LINE, MUTED, OK_BG, OK_FG, TEXT,
                         WARN_BG, WARN_FG, GridTable, Scrolled, clear, open_path, pill)

WRAP = 820          # ширина текста на странице
STRIPE = {"changed": "#E0A800", "only_a": "#C0504D", "only_b": "#4F81BD"}


class ResultView(tk.Frame):
    """Показывает одну сверку из истории. app - главное окно (хранилище, шрифты, переходы)."""

    def __init__(self, parent, app):
        super().__init__(parent, bg=BG)
        self.app, self.fonts = app, app.fonts
        self.run = None
        self.summary = {}
        self.month, self.page = "", 1
        self.view = tk.StringVar(value="cards")
        self.only_diff = tk.BooleanVar()
        self.checks_open = False

        self.area = Scrolled(self)
        self.area.pack(fill="both", expand=True)
        self.page_box = tk.Frame(self.area.body, bg=BG, padx=16, pady=12)
        self.page_box.pack(fill="both", expand=True)

        bar = ttk.Frame(self, padding=(12, 8))
        bar.pack(fill="x")
        self.save_button = ttk.Button(bar, text="Сохранить Excel-отчёт…", command=self.save)
        self.open_button = ttk.Button(bar, text="Открыть отчёт", command=self.open_report)
        self.again_button = ttk.Button(bar, text="Сверить заново", command=self.again)
        self.delete_button = ttk.Button(bar, text="Удалить сверку", command=self.delete)
        for button in (self.save_button, self.open_button, self.again_button):
            button.pack(side="left", padx=(0, 8))
        self.delete_button.pack(side="right")
        self.status = ttk.Label(bar, text="", style="Hint.TLabel")
        self.status.pack(side="left", padx=8)
        self.show(None)

    # ---------- что показывать ----------

    def show(self, run_id):
        """Открывает сверку по номеру (None - пустая вкладка)."""
        self.run = self.app.store.run(run_id) if run_id else None
        self.summary = json.loads(self.run["summary"] or "{}") if self.run else {}
        self.month, self.page, self.checks_open = "", 1, False
        self.status.configure(text="")
        self.render()
        self.area.to_top()

    def label(self, parent, text, font="base", fg=TEXT, bg=BG, **options):
        options.setdefault("wraplength", WRAP)
        return tk.Label(parent, text=text, font=self.fonts[font], fg=fg, bg=bg, justify="left",
                        anchor="w", **options)

    def heading(self, text):
        self.label(self.page_box, text, "head").pack(fill="x", pady=(16, 6))

    def note(self, text, kind="warn"):
        bg, fg = {"warn": (WARN_BG, WARN_FG), "bad": (BAD_BG, BAD_FG)}[kind]
        self.label(self.page_box, text, fg=fg, bg=bg, padx=12, pady=8).pack(fill="x", pady=(0, 6))

    def render(self):
        clear(self.page_box)
        run = self.run
        state = "normal" if run else "disabled"
        has_report = bool(run) and run["status"] != "error" \
            and self.app.store.report_path(run["id"]).is_file()
        self.save_button.configure(state="normal" if has_report else "disabled")
        self.open_button.configure(state="normal" if has_report else "disabled")
        self.again_button.configure(state=state)
        self.delete_button.configure(state=state)
        if run is None:
            self.label(self.page_box, "Здесь появится результат сверки.", "head").pack(
                fill="x", pady=(30, 6))
            self.label(self.page_box, "Запустите сверку на вкладке «Новая сверка» или откройте "
                       "прошлую на вкладке «История».", fg=MUTED).pack(fill="x")
            return

        a, b = run["side_a"] or "Файл 1", run["side_b"] or "Файл 2"
        self.label(self.page_box, service.run_title(run), "title").pack(fill="x")
        period = (f"период {service.period_text(run)}"
                  + (" (взят из дат в файле)" if run["period_auto"] else "")
                  if run["period_from"] else "период не указан")
        self.label(self.page_box, f"{a} ↔ {b}  ·  {period}  ·  сверено "
                   f"{service.format_moment(run['created_at'])}\nФайлы: "
                   f"{Path(run['file_a']).name}  и  {Path(run['file_b']).name}",
                   "small", MUTED).pack(fill="x", pady=(2, 10))

        if run["status"] == "error":
            self.label(self.page_box, "  Сверка не выполнена", "big", MUTED, GREY_BG,
                       pady=8).pack(fill="x")
            self.label(self.page_box, f"Причина: {run['error_text']}").pack(fill="x", pady=(12, 4))
            self.label(self.page_box, "Проверьте файлы и запустите сверку ещё раз.",
                       fg=MUTED).pack(fill="x")
            return

        bad_rows = run["changed"] + run["only_a"] + run["only_b"]
        if run["status"] == "ok":
            self.label(self.page_box, "  Данные совпадают", "big", OK_FG, OK_BG, pady=8).pack(fill="x")
        else:
            text = "  Есть расхождения" + (f": {service.plural_diff(bad_rows)} в строках"
                                           if bad_rows else "")
            self.label(self.page_box, text, "big", BAD_FG, BAD_BG, pady=8).pack(fill="x")

        metrics = tk.Frame(self.page_box, bg=BG)
        metrics.pack(fill="x", pady=(10, 0))
        figures = ((f"Строк в {run['in_a']}", run["rows_a"]), (f"Строк в {run['in_b']}", run["rows_b"]),
                   ("Совпало полностью", run["exact"]), ("Строк с отличиями", run["changed"]),
                   (f"Только в {run['in_a']}", run["only_a"]), (f"Только в {run['in_b']}", run["only_b"]))
        for k, (title, value) in enumerate(figures):
            box = tk.Frame(metrics, bg=CARD, highlightthickness=1, highlightbackground=LINE,
                           padx=10, pady=6)
            box.grid(row=0, column=k, sticky="nsew", padx=(0, 6))
            metrics.columnconfigure(k, weight=1, uniform="metric")
            self.label(box, title, "small", MUTED, CARD, wraplength=120).pack(anchor="w")
            self.label(box, service.number(value), "head", bg=CARD).pack(anchor="w")

        self.render_months()
        whole = self.summary.get("whole") or []
        structure = self.summary.get("structure") or []
        if structure or whole:
            self.heading("Отличия во всём файле")
            items = [(d["Где"], d["Что отличается"]) for d in structure]
            items += [(f"Столбец {col}", f"Во всех строках: «{v1}» в {run['in_a']}, «{v2}» в "
                       f"{run['in_b']}") for _, col, v1, v2 in whole]
            self.pairs_table(items, BAD_BG)
            if whole:
                self.label(self.page_box, "Столбец, который во всём файле отличается одним "
                           "значением, считается расхождением. Если он служебный, исключите его: "
                           "«Дополнительные параметры» → «Не сравнивать столбцы».",
                           "small", MUTED).pack(fill="x", pady=(4, 0))

        self.rows_box = tk.Frame(self.page_box, bg=BG)
        self.rows_box.pack(fill="x")
        self.render_rows()

        tech = self.summary.get("tech") or []
        if tech:
            self.heading("Не учтено в итоге")
            self.pairs_table([(f"Столбец {col}", f"Отметка времени, одна на весь файл: «{v1}» "
                               f"и «{v2}». Похоже на время формирования отчёта, расхождением "
                               "данных не считается") for _, col, v1, v2 in tech], CARD)
        if self.summary.get("checks"):
            self.checks_button = ttk.Button(self.page_box, command=self.toggle_checks)
            self.checks_button.pack(anchor="w", pady=(16, 4))
            self.checks_box = tk.Frame(self.page_box, bg=BG)
            self.checks_box.pack(fill="x")
            self.render_checks()

    def pairs_table(self, items, bg):
        box = tk.Frame(self.page_box, bg=LINE)
        box.pack(fill="x")
        box.columnconfigure(1, weight=1)
        for k, (left, right) in enumerate(items):
            self.label(box, left, bg=CARD, padx=8, pady=4, wraplength=230).grid(
                row=k, column=0, sticky="nsew", padx=(1, 0), pady=(1, 0 if k < len(items) - 1 else 1))
            self.label(box, right, bg=bg, padx=8, pady=4, wraplength=WRAP - 270).grid(
                row=k, column=1, sticky="nsew", padx=1, pady=(1, 0 if k < len(items) - 1 else 1))

    # ---------- по месяцам ----------

    def render_months(self):
        run = self.run
        self.heading("По месяцам")
        if not run["by_month"]:
            if not run["period_from"]:
                reason = self.summary.get("period_note") or "в файлах не найден столбец с датой."
                self.note("Период не указан, и взять его из файла не удалось: " + reason
                          + " Поэтому разбивки по месяцам нет. Нажмите «Сверить заново» и "
                          "укажите период либо задайте столбец с датой в справочнике отчётов.")
            else:
                self.note("В файлах не найден столбец с датой, поэтому расхождения не разложены "
                          "по месяцам. Столбец можно указать в справочнике отчётов и сверить "
                          "заново.")
            return
        self.chips_box = tk.Frame(self.page_box, bg=BG)
        self.chips_box.pack(fill="x")
        self.render_chips()
        self.label(self.page_box, f"Месяц строки взят из столбца «{run['date_column']}». "
                   "Нажмите на месяц, чтобы увидеть только его расхождения.",
                   "small", MUTED).pack(fill="x", pady=(4, 0))

    def render_chips(self):
        clear(self.chips_box)
        chips = [("", "Все", "plain")]
        for m in self.app.store.months(self.run["id"]):
            text = f"{service.month_title(m['month'])}: " + (str(m["bad_rows"]) if m["bad_rows"]
                                                              else "совпало")
            chips.append((m["month"], text + (" · часть месяца" if m["partial"] else ""),
                          "bad" if m["bad_rows"] else "ok"))
        if self.run["unplaced"]:
            chips.append(("none", f"Без даты или вне периода: {self.run['unplaced']}", "bad"))
        for k, (key, text, kind) in enumerate(chips):
            chip = pill(self.chips_box, text, kind, self.fonts, cursor="hand2", pady=5,
                        highlightthickness=2,
                        highlightbackground=TEXT if key == self.month else BG)
            chip.grid(row=k // 5, column=k % 5, sticky="w", padx=(0, 6), pady=3)
            chip.bind("<Button-1>", lambda e, key=key: self.choose_month(key))

    def choose_month(self, key):
        self.month, self.page = key, 1
        self.render_chips()
        self.render_rows()

    # ---------- расхождения в строках ----------

    def render_rows(self):
        clear(self.rows_box)
        run = self.run
        stored = self.app.store.diffs(run["id"], "", 0, 1)[1]
        if not stored:
            return
        groups, total, pages = service.load_page(self.app.store, run, self.month, self.page)
        self.label(self.rows_box, "Расхождения в строках", "head").pack(fill="x", pady=(16, 6))
        top = tk.Frame(self.rows_box, bg=BG)
        top.pack(fill="x", pady=(0, 6))
        bad_rows = run["changed"] + run["only_a"] + run["only_b"]
        first = (self.page - 1) * service.PAGE + 1
        info = f"Показано {first}–{min(self.page * service.PAGE, total)} из {total}." if total else ""
        if stored < bad_rows:
            info += f" В окне хранятся первые {stored} строк, все — в Excel-отчёте."
        self.label(top, info, fg=MUTED, wraplength=380).pack(side="left")
        if self.view.get() == "table":
            ttk.Checkbutton(top, text="Только столбцы с отличиями", variable=self.only_diff,
                            command=self.render_rows).pack(side="right")
        for value, text in (("table", "таблица"), ("cards", "карточки")):
            ttk.Radiobutton(top, text=text, value=value, variable=self.view,
                            command=self.render_rows).pack(side="right", padx=(0, 10))
        self.label(top, "Вид:", fg=MUTED).pack(side="right", padx=(0, 6))

        many_sheets = len(self.summary.get("tables", [])) > 1
        if not groups:
            self.label(self.rows_box, "За этот месяц расхождений в строках нет.",
                       fg=MUTED).pack(fill="x", pady=8)
        for group in groups:
            if many_sheets:
                self.label(self.rows_box, f"Лист «{group['sheet']}»", "bold").pack(fill="x", pady=(8, 4))
            if self.view.get() == "table":
                self.table(group)
            else:
                for row in group["rows"]:
                    self.card(row)
        if self.view.get() == "table" and groups:
            legend = tk.Frame(self.rows_box, bg=BG)
            legend.pack(fill="x", pady=(6, 0))
            for text, bg, note in (("значение", "#FFD966", "в файлах разное"),
                                   ("значение", WARN_BG, "есть только в этом файле"),
                                   ("(пусто)", BAD_BG, "не заполнено"),
                                   ("значение", "#F8CBAD", "строки нет в другом файле")):
                tk.Label(legend, text=text, bg=bg, font=self.fonts["small"], padx=6).pack(side="left")
                tk.Label(legend, text=note + "   ", bg=BG, fg=MUTED,
                         font=self.fonts["small"]).pack(side="left")
        if pages > 1:
            pager = tk.Frame(self.rows_box, bg=BG)
            pager.pack(fill="x", pady=10)
            ttk.Button(pager, text="← Назад", command=lambda: self.turn(-1),
                       state="normal" if self.page > 1 else "disabled").pack(side="left")
            self.label(pager, f"Страница {self.page} из {pages}", fg=MUTED).pack(side="left", padx=12)
            ttk.Button(pager, text="Дальше →", command=lambda: self.turn(1),
                       state="normal" if self.page < pages else "disabled").pack(side="left")

    def turn(self, step):
        self.page += step
        self.render_rows()
        self.area.canvas.update_idletasks()
        self.area.canvas.yview_moveto(max(0, self.rows_box.winfo_y() - 10)
                                      / max(1, self.area.body.winfo_height()))

    def row_title(self, row):
        run = self.run
        if row["kind"] == "changed":
            if row["row_a"] == row["row_b"]:
                return f"Строка {row['row_a']} в {run['in_a']} и в {run['in_b']}"
            return f"Строка {row['row_a']} в {run['in_a']} · строка {row['row_b']} в {run['in_b']}"
        if row["kind"] == "only_a":
            return f"Строка {row['row_a']} в {run['in_a']}"
        return f"Строка {row['row_b']} в {run['in_b']}"

    def card(self, row):
        """Карточка: одна запись с расхождением, в ней только отличающиеся поля."""
        run = self.run
        outer = tk.Frame(self.rows_box, bg=STRIPE[row["kind"]])
        outer.pack(fill="x", pady=(0, 8))
        box = tk.Frame(outer, bg=CARD, highlightthickness=1, highlightbackground=LINE, padx=12, pady=8)
        box.pack(fill="x", padx=(4, 0))
        head = tk.Frame(box, bg=CARD)
        head.pack(fill="x")
        self.label(head, self.row_title(row), "bold", bg=CARD).pack(side="left")
        if row["kind"] == "changed":
            pill(head, row["count"], "warn", self.fonts).pack(side="left", padx=8)
        else:
            missing = run["in_b"] if row["kind"] == "only_a" else run["in_a"]
            pill(head, f"нет в {missing}", "bad", self.fonts).pack(side="left", padx=8)
        if row["month"]:
            pill(head, service.month_title(row["month"]), "plain", self.fonts).pack(side="left")
        if row["ident"]:
            ident = "     ".join(f"{name}: {value}" for (_, name), value in row["ident"])
            self.label(box, ident, fg=MUTED, bg=CARD, wraplength=WRAP - 60).pack(fill="x", pady=(4, 0))
        if row["kind"] == "changed":
            titles = ("Поле", run["side_a"], run["side_b"], "Что не так")
            self.fields(box, titles, [(label, va, vb, note, True)
                                      for label, va, vb, note in row["diffs"]], strong=True)
        full = tk.Frame(box, bg=CARD)
        button = ttk.Button(box, text=f"▸ Вся строка: {row['fields']}", style="Link.TButton")
        button.configure(command=lambda: self.toggle_full(row, full, button))
        button.pack(anchor="w", pady=(6, 0))

    def toggle_full(self, row, full, button):
        if full.winfo_ismapped():
            full.pack_forget()
            button.configure(text=f"▸ Вся строка: {row['fields']}")
            return
        if not full.winfo_children():
            if row["kind"] == "changed":
                titles = ("Поле", self.run["side_a"], self.run["side_b"])
                self.fields(full, titles, [(label, va, vb, None, differs)
                                           for label, va, vb, differs in row["full"]])
            else:
                self.fields(full, ("Поле", "Значение"),
                            [(label, va, None, None, False) for label, va, _, _ in row["full"]],
                            single=True)
        full.pack(fill="x", pady=(4, 0))
        button.configure(text=f"▾ Вся строка: {row['fields']}")

    def fields(self, parent, titles, items, strong=False, single=False):
        """Таблица полей записи: название, значение в файле 1, в файле 2, что не так."""
        grid = tk.Frame(parent, bg=LINE)
        grid.pack(fill="x", pady=(6, 0))
        widths = (220, 520) if single else (200, 215, 215, 150) if len(titles) == 4 else (220, 260, 260)
        for c, title in enumerate(titles):
            self.label(grid, title, "small", MUTED, GREY_BG, padx=8, pady=3,
                       wraplength=widths[c]).grid(row=0, column=c, sticky="nsew", padx=(1, 0), pady=(1, 0))
            grid.columnconfigure(c, weight=1 if c else 0, minsize=widths[c] + 16)
        for r, ((letter, name), va, vb, note, differs) in enumerate(items, 1):
            cells = [(f"{letter}  {name}" if letter else name, TEXT, CARD, "base")]
            for value in ((va,) if single else (va, vb)):
                if value is None:
                    cells.append(("пусто", BAD_FG if differs else "#A6A6A6",
                                  BAD_BG if differs and strong else CARD, "italic"))
                else:
                    cells.append((value, TEXT, WARN_BG if differs else CARD,
                                  "bold" if differs else "base"))
            if note is not None:
                cells.append((note, MUTED, CARD, "base"))
            for c, (text, fg, bg, font) in enumerate(cells):
                self.label(grid, text, font, fg, bg, padx=8, pady=3, wraplength=widths[c]).grid(
                    row=r, column=c, sticky="nsew", padx=(1, 1 if c == len(cells) - 1 else 0),
                    pady=(1, 1 if r == len(items) else 0))

    def table(self, group):
        """Широкая таблица, как в Excel: по две строки на запись, отличия подсвечены."""
        run = self.run
        columns, first_tech, first_whole = group["columns"], group["first_tech"], group["first_whole"]
        shown = list(range(len(columns)))
        if self.only_diff.get():
            differing = {int(k) for row in group["rows"] for k in row["kinds"]}
            only_rows = any(row["kind"] != "changed" for row in group["rows"])
            shown = [i for i in shown if i in differing or first_whole <= i < first_tech
                     or (only_rows and not differing)]
        headers = ["Файл", "Строка"] + [columns[i] + ("\nтехнический" if i >= first_tech else "")
                                        for i in shown] + ["Что проверить"]
        rows = []
        for row in group["rows"]:
            sides = [(run["side_a"], row["row_a"], row["a"]), (run["side_b"], row["row_b"], row["b"])]
            sides = [s for s in sides if s[2] is not None]
            for n, (side, number, values) in enumerate(sides):
                cells = [(side, "side", 1), (str(number), "plain", 1)]
                for i in shown:
                    value = values[i] if i < len(values) else None
                    kind = row["kinds"].get(str(i))
                    if i >= first_tech:
                        cells.append((value or "", "tech", 1))
                    elif row["kind"] != "changed":
                        cells.append((value or "", row["kind"], 1))
                    elif kind and value is None:
                        cells.append(("(пусто)", "e", 1))
                    elif kind:
                        cells.append((value, "have" if kind in ("empty_a", "empty_b") else "x", 1))
                    elif i >= first_whole:
                        cells.append((value or "", "x", 1))
                    else:
                        cells.append((value or "", "same", 1))
                cells.append((row["hint"] or "", "plain", len(sides)) if n == 0 else None)
                rows.append(cells)
        table = GridTable(self.rows_box, self.fonts)
        table.pack(fill="x", pady=(0, 6))
        table.show(headers, rows)

    # ---------- проверки ----------

    def toggle_checks(self):
        self.checks_open = not self.checks_open
        self.render_checks()

    def render_checks(self):
        clear(self.checks_box)
        self.checks_button.configure(
            text=("▾" if self.checks_open else "▸") + " Все проверки и замечания")
        if not self.checks_open:
            return
        many_sheets = len(self.summary.get("tables", [])) > 1
        grid = tk.Frame(self.checks_box, bg=LINE)
        grid.pack(fill="x")
        widths = (230, 130, 90, 330)
        colors = {"совпадает": OK_BG, "есть различия": BAD_BG, "обратить внимание": WARN_BG}
        checks = self.summary["checks"]
        for r, check in enumerate(checks):
            name = (f"[{check['Лист']}] " if many_sheets and check["Лист"] else "") + check["Проверка"]
            values = f"{check['a']} / {check['b']}" if check["a"] != "" or check["b"] != "" else ""
            row = ((name, CARD), (check["Результат"], colors.get(check["Результат"], GREY_BG)),
                   (values, CARD), (check["Комментарий"], CARD))
            for c, (text, bg) in enumerate(row):
                self.label(grid, text, bg=bg, padx=8, pady=3, wraplength=widths[c]).grid(
                    row=r, column=c, sticky="nsew", padx=(1, 1 if c == 3 else 0),
                    pady=(1, 1 if r == len(checks) - 1 else 0))
            grid.columnconfigure(3, weight=1)

    # ---------- кнопки ----------

    def save(self):
        source = self.app.store.report_path(self.run["id"])
        path = filedialog.asksaveasfilename(
            parent=self, title="Сохранить отчёт", initialdir=self.app.last_dir,
            initialfile=service.report_file_name(self.run), defaultextension=".xlsx",
            filetypes=[("Файлы Excel", "*.xlsx")])
        if not path:
            return
        try:
            shutil.copyfile(source, path)
        except PermissionError:
            messagebox.showerror(self.app.title(), f"Нет доступа к файлу {path}. Если он открыт "
                                 "в Excel, закройте его и сохраните отчёт снова", parent=self)
            return
        except OSError as e:
            messagebox.showerror(self.app.title(), f"Не удалось сохранить отчёт: {e}", parent=self)
            return
        self.app.last_dir = str(Path(path).parent)
        self.status.configure(text=f"Отчёт сохранён: {path}")

    def open_report(self):
        try:
            open_path(self.app.store.report_path(self.run["id"]))
        except OSError as e:
            messagebox.showerror(self.app.title(), f"Не удалось открыть отчёт: {e}", parent=self)

    def again(self):
        self.app.prefill(self.run)

    def delete(self):
        if not messagebox.askyesno(self.app.title(), "Удалить эту сверку из истории вместе "
                                   "с её Excel-отчётом?\n\nВаши исходные файлы останутся на месте.",
                                   parent=self):
            return
        self.app.store.delete_run(self.run["id"])
        self.show(None)
        self.app.history.refresh()
        self.app.dicts.refresh()
        self.app.tabs.select(self.app.history)
