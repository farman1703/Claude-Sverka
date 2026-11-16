"""Вкладка «История»: все прошлые сверки с отбором по системе, отчёту и итогу."""

import tkinter as tk
from tkinter import messagebox, ttk

import service
from gui_widgets import BG, open_path

ANY_SYSTEM, ANY_REPORT = "Все системы", "Все отчёты"
STATUSES = (("Любой итог", None), ("Совпало", "ok"), ("Есть расхождения", "diff"),
            ("Ошибка", "error"))
ROW_COLORS = {"ok": "#E8F6E3", "bad": "#FBE3E3", "err": "#EFEEEA"}


class HistoryView(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG, padx=14, pady=12)
        self.app = app
        self.system, self.report = tk.StringVar(value=ANY_SYSTEM), tk.StringVar(value=ANY_REPORT)
        self.status = tk.StringVar(value=STATUSES[0][0])
        self.system_ids, self.report_ids = {}, {}

        filters = ttk.Frame(self)
        filters.pack(fill="x", pady=(0, 8))
        self.system_box = ttk.Combobox(filters, textvariable=self.system, state="readonly", width=24)
        self.report_box = ttk.Combobox(filters, textvariable=self.report, state="readonly", width=36)
        status_box = ttk.Combobox(filters, textvariable=self.status, state="readonly", width=18,
                                  values=[s[0] for s in STATUSES])
        for box in (self.system_box, self.report_box, status_box):
            box.pack(side="left", padx=(0, 8))
        self.system_box.bind("<<ComboboxSelected>>", lambda e: self.refresh(system_changed=True))
        self.report_box.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        status_box.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        self.found = ttk.Label(filters, text="", style="Hint.TLabel")
        self.found.pack(side="left", padx=8)

        table = ttk.Frame(self)
        table.pack(fill="both", expand=True)
        columns = (("when", "Когда", 130), ("system", "Система", 150), ("report", "Отчёт", 250),
                   ("period", "Период", 190), ("result", "Итог", 150))
        self.tree = ttk.Treeview(table, columns=[c[0] for c in columns], show="headings",
                                 selectmode="browse")
        for key, title, width in columns:
            self.tree.heading(key, text=title, anchor="w")
            self.tree.column(key, width=width, anchor="w", stretch=key == "report")
        for kind, color in ROW_COLORS.items():
            self.tree.tag_configure(kind, background=color)
        bar = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", lambda e: self.open())
        self.tree.bind("<Return>", lambda e: self.open())
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.update_buttons())

        buttons = ttk.Frame(self)
        buttons.pack(fill="x", pady=(8, 0))
        self.open_button = ttk.Button(buttons, text="Открыть сверку", command=self.open)
        self.delete_button = ttk.Button(buttons, text="Удалить сверку", command=self.delete)
        self.open_button.pack(side="left")
        self.delete_button.pack(side="left", padx=8)
        ttk.Button(buttons, text="Папка с данными", command=self.open_folder).pack(side="right")
        ttk.Label(buttons, style="Hint.TLabel", text="Открыть сверку можно двойным щелчком."
                  ).pack(side="left", padx=8)
        self.refresh()

    def selected(self):
        chosen = self.tree.selection()
        return int(chosen[0]) if chosen else None

    def update_buttons(self):
        state = "normal" if self.selected() else "disabled"
        self.open_button.configure(state=state)
        self.delete_button.configure(state=state)

    def refresh(self, system_changed=False):
        """Перечитывает справочники для отбора и список сверок."""
        store = self.app.store
        systems = store.systems()
        self.system_ids = {s["name"]: s["id"] for s in systems}
        self.system_box.configure(values=[ANY_SYSTEM] + list(self.system_ids))
        if self.system.get() not in self.system_ids:
            self.system.set(ANY_SYSTEM)
        system_id = self.system_ids.get(self.system.get())
        # В списке отчётов - отчёты выбранной системы; без системы - все, с её названием
        self.report_ids = {(r["name"] if system_id else f"{r['system']} / {r['name']}"): r["id"]
                           for r in store.reports(system_id)}
        self.report_box.configure(values=[ANY_REPORT] + list(self.report_ids))
        if system_changed or self.report.get() not in self.report_ids:
            self.report.set(ANY_REPORT)
        status = dict(STATUSES)[self.status.get()]
        runs, total = store.runs(system_id, self.report_ids.get(self.report.get()), status)

        self.tree.delete(*self.tree.get_children())
        for run in runs:
            kind, label = service.run_label(run)
            self.tree.insert("", "end", iid=str(run["id"]), tags=(kind,), values=(
                service.format_moment(run["created_at"]), run["system"] or "—",
                run["report"] or run["title"] or "Сверка двух файлов",
                service.period_text(run), label))
        shown = f"Найдено: {total}" + (f", показаны последние {len(runs)}" if total > len(runs) else "")
        if not total:
            filtered = system_id or status or self.report.get() != ANY_REPORT
            shown = "По этим условиям сверок нет." if filtered else "Сверок пока не было."
        self.found.configure(text=shown)
        self.update_buttons()

    def open(self):
        run_id = self.selected()
        if run_id:
            self.app.show_run(run_id)

    def delete(self):
        run_id = self.selected()
        if not run_id or not messagebox.askyesno(
                self.app.title(), "Удалить эту сверку из истории вместе с её Excel-отчётом?\n\n"
                "Ваши исходные файлы останутся на месте.", parent=self):
            return
        self.app.store.delete_run(run_id)
        if self.app.result.run is not None and self.app.result.run["id"] == run_id:
            self.app.result.show(None)
        self.refresh()
        self.app.dicts.refresh()

    def open_folder(self):
        try:
            open_path(self.app.store.folder)
        except OSError as e:
            messagebox.showerror(self.app.title(), f"Не удалось открыть папку: {e}", parent=self)
