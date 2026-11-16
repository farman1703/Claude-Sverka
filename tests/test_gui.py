#!/usr/bin/env python3
"""
Проверка окна: весь путь пользователя от справочников до истории.

Нужен экран (в Linux без экрана: xvfb-run python tests/test_gui.py).
Если задать переменную SHOTS=папка, туда сохраняются снимки вкладок
(нужна программа import из ImageMagick).
"""

import os
import subprocess
import sys
import tempfile
import time
import tkinter as tk
import unittest
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import Workbook, load_workbook  # noqa: E402

import sverka_gui  # noqa: E402
from storage import Store  # noqa: E402

HEAD = ["ID", "Дата", "Сумма", "Клиент", "Сформировано"]
SHOTS = os.environ.get("SHOTS")


def rows(stamp, count=70):
    return [[f"a{k:03d}", datetime(2026, 7 + k % 3, 1 + k % 27), 100 + k, f"Клиент {k}", stamp]
            for k in range(count)]


def save(path, data):
    book = Workbook()
    book.active.append(HEAD)
    for r in data:
        book.active.append(r)
    book.save(path)
    return path


class Window(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            tk.Tk().destroy()
        except tk.TclError:
            raise unittest.SkipTest("нет экрана") from None

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self.tmp.name)
        self.shown = []
        messagebox.showerror = lambda t, m, **k: self.shown.append(m)
        messagebox.showwarning = lambda t, m, **k: self.shown.append(m)
        messagebox.askyesno = lambda *a, **k: True
        self.app = sverka_gui.App(Store(self.folder / "data"))
        self.app.update()

    def tearDown(self):
        self.app.close()
        self.tmp.cleanup()

    def shot(self, name):
        if SHOTS:
            self.app.update()
            time.sleep(0.2)
            self.app.update()
            subprocess.run(["import", "-window", "root", str(Path(SHOTS) / f"{name}.png")])

    def run_compare(self):
        """Нажимает «Сверить» и ждёт, пока результат появится в окне."""
        before = self.app.store.runs()[1]
        self.app.result.show(None)
        self.app.run_button.invoke()
        end = time.time() + 120
        while time.time() < end and not self.shown:
            self.app.update()
            done = self.app.store.runs()[1] > before and self.app.result.run is not None
            if done and str(self.app.run_button.cget("state")) == "normal":
                break
            time.sleep(0.05)
        self.app.update()
        return self.app.result.run

    def test_full_path(self):
        app = self.app
        base = rows("13.11.2026 10:00:00")
        other = [list(r) for r in rows("13.11.2026 10:07:30")]
        for k in range(0, 60):
            other[k][2] += 1                      # 60 строк с другой суммой
        other[61][3] = None                       # пустая ячейка
        del other[65]                             # строки нет во втором файле
        file1 = save(self.folder / "продажи_dk_20261113.xlsx", base)
        file2 = save(self.folder / "продажи_ng_20261113.xlsx", other)

        # --- справочники ---
        app.tabs.select(app.dicts)
        app.dicts.new_systems.insert("1.0", "CRM\nБиллинг\nCRM")
        app.dicts.add_systems()
        self.assertEqual(len(app.dicts.systems.get_children()), 2)
        app.dicts.systems.selection_set(app.dicts.systems.get_children()[0])       # CRM
        app.update()
        app.dicts.new_reports.insert("1.0", "Продажи\nОстатки\nКлиенты")
        app.dicts.add_reports()
        self.assertEqual(len(app.dicts.reports.get_children()), 3)
        self.assertIn("Добавлено: 3", app.dicts.message.cget("text"))
        app.dicts.search.set("ост")
        self.assertEqual(len(app.dicts.reports.get_children()), 1)
        app.dicts.search.set("")
        self.shot("dicts")

        # --- новая сверка: проверки ввода ---
        app.tabs.select(app.new_tab)
        self.assertEqual(list(app.system_box.cget("values")), [sverka_gui.NO_SYSTEM, "CRM", "Биллинг"])
        app.run_button.invoke()
        self.assertEqual(self.shown.pop(), "Файл 1 не выбран")
        app.file1.set(str(file1))
        app.file2.set(f'"{file2}"')
        app.system.set("CRM")
        app.fill_reports()
        self.assertEqual(list(app.report_box.cget("values")), ["Клиенты", "Остатки", "Продажи"])
        app.run_button.invoke()
        self.assertEqual(self.shown.pop(), "Выберите отчёт.")
        app.report.set("Продажи")
        app.date_from.set("01.07.2026")
        app.run_button.invoke()
        self.assertIn("обе даты", self.shown.pop())
        app.date_to.set("вчера")
        app.run_button.invoke()
        self.assertIn("Не понял дату", self.shown.pop())
        app.set_period("month")
        self.assertTrue(app.date_from.get().startswith("01."))
        app.set_period("clear")
        self.assertEqual(app.store.runs()[1], 0)
        self.shot("new")

        # --- сверка ---
        run = self.run_compare()
        self.assertEqual(self.shown, [])
        self.assertEqual((run["status"], run["report"], run["system"]), ("diff", "Продажи", "CRM"))
        self.assertEqual((run["changed"], run["only_a"], run["period_auto"], run["by_month"]),
                         (61, 1, 1, 1))
        self.assertEqual(app.tabs.select(), str(app.result))
        self.assertEqual(str(app.result.save_button.cget("state")), "normal")
        self.shot("result")

        result = app.result
        result.choose_month("2026-08")
        self.assertEqual(result.month, "2026-08")
        result.choose_month("")
        result.turn(1)
        self.assertEqual(result.page, 2)
        result.turn(-1)
        result.view.set("table")
        result.render_rows()
        self.shot("result_table")
        result.only_diff.set(True)
        result.render_rows()
        result.view.set("cards")
        result.render_rows()
        result.toggle_checks()
        self.assertTrue(result.checks_open)
        app.update()

        # --- Excel-отчёт ---
        target = self.folder / "сохранённый отчёт.xlsx"
        filedialog.asksaveasfilename = lambda **k: str(target)
        result.save_button.invoke()
        self.assertIn("Продажи, CRM: dk ↔ ng", load_workbook(target)["Итог"]["A1"].value)
        self.assertIn("Отчёт сохранён", result.status.cget("text"))

        # --- сверить заново: форма заполнена из сверки ---
        result.again_button.invoke()
        self.assertEqual(app.tabs.select(), str(app.new_tab))
        self.assertEqual((app.system.get(), app.report.get(), app.file1.get(), app.date_from.get()),
                         ("CRM", "Продажи", str(file1), ""))

        # --- сверка без справочника и сверка с ошибкой ---
        app.system.set(sverka_gui.NO_SYSTEM)
        app.fill_reports()
        app.file2.set(str(file1))
        quick = self.run_compare()
        self.assertEqual((quick["status"], quick["report"], quick["side_a"]), ("ok", None, "Файл 1"))
        broken = self.folder / "broken.xlsx"
        broken.write_text("не Excel", encoding="utf-8")
        app.file2.set(str(broken))
        failed = self.run_compare()
        self.assertEqual(failed["status"], "error")
        self.assertEqual(str(app.result.save_button.cget("state")), "disabled")
        self.shot("result_error")

        # --- история ---
        app.tabs.select(app.history)
        history = app.history
        history.refresh()
        self.assertEqual(len(history.tree.get_children()), 3)
        self.assertIn("Найдено: 3", history.found.cget("text"))
        self.shot("history")
        history.status.set("Есть расхождения")
        history.refresh()
        self.assertEqual(history.tree.get_children(), (str(run["id"]),))
        history.status.set("Любой итог")
        history.system.set("CRM")
        history.refresh(system_changed=True)
        self.assertEqual(len(history.tree.get_children()), 1)
        self.assertEqual(list(history.report_box.cget("values"))[1:], ["Клиенты", "Остатки", "Продажи"])
        history.system.set("Все системы")
        history.refresh(system_changed=True)
        history.tree.selection_set(str(run["id"]))
        history.open()
        self.assertEqual((app.result.run["id"], app.tabs.select()), (run["id"], str(app.result)))

        # --- удаление: отчёт со сверками удалить нельзя, после удаления сверки можно ---
        app.dicts.refresh()
        report_id = [r for r in app.dicts.reports.get_children()
                     if app.dicts.reports.item(r, "values")[0] == "Продажи"][0]
        self.assertEqual(app.dicts.reports.item(report_id, "values")[1:],
                         ("Дата", "найден программой", "1"))
        app.dicts.reports.selection_set(report_id)
        app.dicts.delete_report()
        self.assertIn("пока по нему есть сверки", self.shown.pop())
        app.result.delete_button.invoke()
        self.assertIsNone(app.result.run)
        self.assertEqual(app.store.runs()[1], 2)
        self.assertFalse(app.store.report_path(run["id"]).exists())
        self.assertTrue(file1.is_file() and file2.is_file())     # исходные файлы на месте
        app.dicts.reports.selection_set(report_id)
        app.dicts.delete_report()
        self.assertEqual(len(app.dicts.reports.get_children()), 2)
        self.assertEqual(self.shown, [])


if __name__ == "__main__":
    unittest.main(verbosity=1)
