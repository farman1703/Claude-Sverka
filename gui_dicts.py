"""Вкладка «Справочники»: системы и отчёты каждой системы."""

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from gui_widgets import BG


class ReportDialog(tk.Toplevel):
    """Окно правки отчёта: название и столбец с датой."""

    def __init__(self, parent, report):
        super().__init__(parent)
        self.title("Отчёт")
        self.transient(parent)
        self.resizable(False, False)
        self.result = None
        self.name = tk.StringVar(value=report["name"])
        self.column = tk.StringVar(value=report["date_column"] or "")
        box = ttk.Frame(self, padding=14)
        box.pack(fill="both", expand=True)
        ttk.Label(box, text="Название:").grid(row=0, column=0, sticky="w", pady=4)
        entry = ttk.Entry(box, textvariable=self.name, width=46)
        entry.grid(row=0, column=1, padx=(8, 0), pady=4)
        ttk.Label(box, text="Столбец с датой:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(box, textvariable=self.column, width=46).grid(row=1, column=1, padx=(8, 0), pady=4)
        ttk.Label(box, style="Hint.TLabel", wraplength=420, justify="left", text=(
            "По этому столбцу расхождения раскладываются по месяцам. Оставьте поле пустым — "
            "программа найдёт столбец сама при первой сверке.")).grid(
                row=2, column=0, columnspan=2, sticky="w", pady=(4, 10))
        buttons = ttk.Frame(box)
        buttons.grid(row=3, column=0, columnspan=2, sticky="e")
        ttk.Button(buttons, text="Сохранить", command=self.accept).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="Отмена", command=self.destroy).pack(side="left")
        self.bind("<Return>", lambda e: self.accept())
        self.bind("<Escape>", lambda e: self.destroy())
        entry.focus_set()
        self.update_idletasks()
        self.geometry(f"+{parent.winfo_rootx() + 80}+{parent.winfo_rooty() + 80}")
        self.grab_set()

    def accept(self):
        self.result = (self.name.get(), self.column.get())
        self.destroy()


class DictView(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG, padx=14, pady=12)
        self.app = app
        self.search = tk.StringVar()
        self.search.trace_add("write", lambda *_: self.fill_reports())
        ttk.Label(self, style="Hint.TLabel", wraplength=860, justify="left", text=(
            "Система — это источник данных, у каждой системы свой список отчётов. Справочники "
            "нужны, чтобы в истории было видно, какой отчёт какой системы сверялся. "
            "Заполнять их не обязательно: сверять файлы можно и без них.")).pack(fill="x", pady=(0, 8))
        panes = ttk.Frame(self)
        panes.pack(fill="both", expand=True)
        panes.columnconfigure(0, weight=2, uniform="pane")
        panes.columnconfigure(1, weight=5, uniform="pane")
        panes.rowconfigure(0, weight=1)

        # --- системы ---
        left = ttk.LabelFrame(panes, text=" Системы ", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.systems = ttk.Treeview(left, columns=("name", "used"), show="headings", height=6,
                                    selectmode="browse")
        self.systems.heading("name", text="Название", anchor="w")
        self.systems.heading("used", text="Отчётов", anchor="w")
        self.systems.column("name", width=150, anchor="w")
        self.systems.column("used", width=70, anchor="w", stretch=False)
        self.systems.pack(fill="both", expand=True)
        self.systems.bind("<<TreeviewSelect>>", lambda e: self.fill_reports())
        row = ttk.Frame(left)
        row.pack(fill="x", pady=6)
        self.rename_button = ttk.Button(row, text="Переименовать…", command=self.rename_system)
        self.rename_button.pack(side="left")
        self.delete_system_button = ttk.Button(row, text="Удалить", command=self.delete_system)
        self.delete_system_button.pack(side="left", padx=6)
        self.new_systems = tk.Text(left, height=3, width=20, wrap="none", relief="solid", borderwidth=1)
        self.new_systems.pack(fill="x")
        ttk.Button(left, text="Добавить системы", command=self.add_systems).pack(anchor="w", pady=(6, 2))
        ttk.Label(left, style="Hint.TLabel", wraplength=230, justify="left",
                  text="Можно сразу несколько — каждая с новой строки.").pack(anchor="w")

        # --- отчёты выбранной системы ---
        right = ttk.LabelFrame(panes, text=" Отчёты ", padding=10)
        right.grid(row=0, column=1, sticky="nsew")
        top = ttk.Frame(right)
        top.pack(fill="x", pady=(0, 6))
        self.reports_title = ttk.Label(top, text="")
        self.reports_title.pack(side="left")
        ttk.Entry(top, textvariable=self.search, width=22).pack(side="right")
        ttk.Label(top, text="Поиск:", style="Hint.TLabel").pack(side="right", padx=6)
        columns = (("name", "Название", 190), ("column", "Столбец с датой", 130),
                   ("source", "Как определён", 190), ("used", "Сверок", 70))
        self.reports = ttk.Treeview(right, columns=[c[0] for c in columns], show="headings",
                                    height=6, selectmode="browse")
        for key, title, width in columns:
            self.reports.heading(key, text=title, anchor="w")
            self.reports.column(key, width=width, anchor="w", stretch=key == "name")
        self.reports.pack(fill="both", expand=True)
        self.reports.bind("<<TreeviewSelect>>", lambda e: self.update_buttons())
        self.reports.bind("<Double-1>", lambda e: self.edit_report())
        row = ttk.Frame(right)
        row.pack(fill="x", pady=6)
        self.edit_button = ttk.Button(row, text="Изменить…", command=self.edit_report)
        self.edit_button.pack(side="left")
        self.delete_report_button = ttk.Button(row, text="Удалить", command=self.delete_report)
        self.delete_report_button.pack(side="left", padx=6)
        self.new_reports = tk.Text(right, height=3, width=40, wrap="none", relief="solid", borderwidth=1)
        self.new_reports.pack(fill="x")
        self.add_reports_button = ttk.Button(right, text="Добавить отчёты в эту систему",
                                             command=self.add_reports)
        self.add_reports_button.pack(anchor="w", pady=(6, 2))
        ttk.Label(right, style="Hint.TLabel", wraplength=560, justify="left", text=(
            "Можно вставить весь список сразу — каждый отчёт с новой строки. Столбец с датой "
            "нужен, чтобы разложить расхождения по месяцам; программа находит его сама.")
                  ).pack(anchor="w")

        self.message = ttk.Label(self, text="", style="Hint.TLabel")
        self.message.pack(fill="x", pady=(8, 0))
        self.refresh()

    # ---------- показ ----------

    def system_id(self):
        chosen = self.systems.selection()
        return int(chosen[0]) if chosen else None

    def report_id(self):
        chosen = self.reports.selection()
        return int(chosen[0]) if chosen else None

    def refresh(self, select=None):
        """Перечитывает оба списка; select - система, которую выбрать."""
        keep = select or self.system_id()
        rows = self.app.store.systems()
        self.systems.delete(*self.systems.get_children())
        for s in rows:
            self.systems.insert("", "end", iid=str(s["id"]), values=(s["name"], s["used"]))
        ids = [s["id"] for s in rows]
        if ids:
            self.systems.selection_set(str(keep if keep in ids else ids[0]))
        self.fill_reports()

    def fill_reports(self):
        system_id = self.system_id()
        self.reports.delete(*self.reports.get_children())
        if system_id is None:
            self.reports_title.configure(text="Сначала добавьте хотя бы одну систему.")
        else:
            name = self.systems.item(str(system_id), "values")[0]
            rows = self.app.store.reports(system_id)
            self.reports_title.configure(text=f"Система «{name}», отчётов: {len(rows)}")
            query = self.search.get().strip().casefold()
            for r in rows:
                if query and query not in r["name"].casefold():
                    continue
                source = ("указан вами" if r["date_source"] == "manual" else
                          "найден программой" if r["date_column"] else
                          "не найден" if r["used"] else "определится при сверке")
                self.reports.insert("", "end", iid=str(r["id"]), values=(
                    r["name"], r["date_column"] or "", source, r["used"]))
        self.update_buttons()

    def update_buttons(self):
        system = "normal" if self.system_id() else "disabled"
        report = "normal" if self.report_id() else "disabled"
        self.rename_button.configure(state=system)
        self.delete_system_button.configure(state=system)
        self.add_reports_button.configure(state=system)
        self.edit_button.configure(state=report)
        self.delete_report_button.configure(state=report)

    def changed(self, text, select=None):
        """После правки: обновить списки здесь и на остальных вкладках."""
        self.message.configure(text=text)
        self.refresh(select)
        self.app.dictionaries_changed()

    @staticmethod
    def added_text(added, repeated):
        parts = [f"Добавлено: {added}."] if added else []
        parts += [f"«{name}» уже есть в списке." for name in repeated]
        return " ".join(parts) or "Впишите названия — каждое с новой строки."

    # ---------- действия ----------

    def add_systems(self):
        added, repeated = self.app.store.add_systems(self.new_systems.get("1.0", "end"))
        if added:
            self.new_systems.delete("1.0", "end")
        self.changed(self.added_text(added, repeated))

    def add_reports(self):
        system_id = self.system_id()
        added, repeated = self.app.store.add_reports(system_id, self.new_reports.get("1.0", "end"))
        if added:
            self.new_reports.delete("1.0", "end")
        self.changed(self.added_text(added, repeated), system_id)

    def rename_system(self):
        system_id = self.system_id()
        old = self.systems.item(str(system_id), "values")[0]
        name = simpledialog.askstring(self.app.title(), "Название системы:", initialvalue=old,
                                      parent=self)
        if name is None:
            return
        try:
            self.app.store.rename_system(system_id, name)
        except ValueError as e:
            messagebox.showerror(self.app.title(), str(e), parent=self)
            return
        self.changed("Сохранено.", system_id)

    def edit_report(self):
        report_id = self.report_id()
        if not report_id:
            return
        dialog = ReportDialog(self, self.app.store.report(report_id))
        self.wait_window(dialog)
        if dialog.result is None:
            return
        try:
            self.app.store.save_report(report_id, *dialog.result)
        except ValueError as e:
            messagebox.showerror(self.app.title(), str(e), parent=self)
            return
        self.changed("Сохранено.", self.system_id())

    def delete_system(self):
        system_id = self.system_id()
        name = self.systems.item(str(system_id), "values")[0]
        if not messagebox.askyesno(self.app.title(), f"Удалить систему «{name}»?", parent=self):
            return
        try:
            self.app.store.delete_system(system_id)
        except ValueError as e:
            messagebox.showerror(self.app.title(), str(e), parent=self)
            return
        self.changed("Удалено.")

    def delete_report(self):
        report_id = self.report_id()
        name = self.reports.item(str(report_id), "values")[0]
        if not messagebox.askyesno(self.app.title(), f"Удалить отчёт «{name}»?", parent=self):
            return
        try:
            self.app.store.delete_report(report_id)
        except ValueError as e:
            messagebox.showerror(self.app.title(), str(e), parent=self)
            return
        self.changed("Удалено.", self.system_id())
