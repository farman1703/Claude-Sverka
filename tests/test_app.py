#!/usr/bin/env python3
"""
Проверка истории сверок, справочников и разбивки по месяцам (без окна).

Запуск:
    python tests/test_app.py
"""

import json
import sys
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import Workbook, load_workbook  # noqa: E402

import service  # noqa: E402
from engine import ce  # noqa: E402
from storage import Store  # noqa: E402

HEAD = ["ID", "Дата", "Сумма", "Клиент"]
BASE = [["a1", datetime(2026, 7, 5), 10, "Иванов"], ["a2", datetime(2026, 7, 20), 20, "Петров"],
        ["a3", datetime(2026, 8, 3), 30, "Сидоров"], ["a4", datetime(2026, 9, 1), 40, "Орлов"],
        ["a5", datetime(2026, 9, 15), 50, "Волков"], ["a6", datetime(2026, 9, 28), 60, "Зайцев"]]


def changed():
    """Три расхождения: июль - другая сумма, сентябрь - другая сумма и пропавшая строка."""
    rows = [list(r) for r in BASE]
    rows[0][2] = 11
    rows[3][2] = 41
    del rows[5]
    return rows


class Case(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self.tmp.name)
        self.store = Store(self.folder / "data")
        self.store.add_systems("CRM\nБиллинг\n\nCRM")
        crm = [s["id"] for s in self.store.systems() if s["name"] == "CRM"][0]
        self.store.add_reports(crm, "Продажи\nОстатки")
        self.report = [r for r in self.store.reports(crm) if r["name"] == "Продажи"][0]
        self.count = 0

    def tearDown(self):
        self.store.close()
        self.tmp.cleanup()

    def xlsx(self, rows, head=HEAD, name=None):
        self.count += 1
        path = self.folder / (name or f"file{self.count}.xlsx")
        book = Workbook()
        book.active.append(head)
        for r in rows:
            book.active.append(r)
        book.save(path)
        return path

    def compare(self, rows_a, rows_b, period=(date(2026, 7, 1), date(2026, 9, 30)), report="std",
                head=HEAD, names=("выгрузка_dk.xlsx", "выгрузка_ng.xlsx"), **options):
        report = self.store.report(self.report["id"]) if report == "std" else report
        a, b = self.xlsx(rows_a, head, names[0]), self.xlsx(rows_b, head, names[1])
        run_id = self.store.create_run(report, a, b, *period, "{}")
        status = service.execute(self.store, run_id, a, b, report, *period, ce.Options(**options))
        return self.store.run(run_id), status

    def months(self, run):
        return {m["month"]: (m["bad_rows"], m["partial"]) for m in self.store.months(run["id"])}

    # ---------- даты ----------

    def test_dates(self):
        for value, expected in (
                (datetime(2026, 7, 5, 12), date(2026, 7, 5)), ("05.07.2026", date(2026, 7, 5)),
                ("2026-07-05", date(2026, 7, 5)), ("05.07.2026 10:00", date(2026, 7, 5)),
                ("2026-07-05T10:00:00", date(2026, 7, 5)), ("2026-07", date(2026, 7, 1)),
                ("07.2026", date(2026, 7, 1)), ("Иванов", None), (5, None), ("", None)):
            self.assertEqual(service.to_date(value), expected, value)
        self.assertEqual(service.months_between(date(2025, 11, 15), date(2026, 2, 1)),
                         ["2025-11", "2025-12", "2026-01", "2026-02"])
        self.assertEqual(service.shift_month("2026-01", -1), "2025-12")
        self.assertEqual(service.month_bounds("2024-02"), (date(2024, 2, 1), date(2024, 2, 29)))
        self.assertEqual(service.month_title("2026-07"), "Июль 2026")
        self.assertEqual(service.parse_user_date(" 31.12.2026 "), date(2026, 12, 31))
        self.assertIsNone(service.parse_user_date(""))
        self.assertRaises(ValueError, service.parse_user_date, "завтра")

    def test_quick_periods(self):
        today = date(2026, 11, 16)
        self.assertEqual(service.quick_period("month", today), (date(2026, 10, 1), date(2026, 10, 31)))
        self.assertEqual(service.quick_period("quarter", today), (date(2026, 7, 1), date(2026, 9, 30)))
        self.assertEqual(service.quick_period("year", today), (date(2026, 1, 1), today))
        january = date(2026, 1, 15)
        self.assertEqual(service.quick_period("month", january), (date(2025, 12, 1), date(2025, 12, 31)))
        self.assertEqual(service.quick_period("quarter", january), (date(2025, 10, 1), date(2025, 12, 31)))

    # ---------- справочники ----------

    def test_dictionaries(self):
        self.assertEqual([(s["name"], s["used"]) for s in self.store.systems()],
                         [("CRM", 2), ("Биллинг", 0)])
        self.assertEqual(self.store.add_systems("CRM\nСклад"), (1, ["CRM"]))
        crm, billing = (s["id"] for s in self.store.systems() if s["name"] in ("CRM", "Биллинг"))
        self.assertEqual(self.store.add_reports(billing, "Продажи\nКлиенты"), (2, []))
        self.assertEqual(self.store.add_reports(crm, "Продажи"), (0, ["Продажи"]))
        self.assertEqual([r["name"] for r in self.store.reports(billing)], ["Клиенты", "Продажи"])
        self.assertEqual(len(self.store.reports()), 4)
        self.store.rename_system(billing, "Биллинг 2")
        self.assertRaises(ValueError, self.store.rename_system, billing, "CRM")
        self.assertRaises(ValueError, self.store.rename_system, billing, "  ")
        # столбец с датой, вписанный вручную, помечается как ручной; очищенный - снова авто
        self.store.save_report(self.report["id"], "Продажи 2", " Дата операции ")
        row = self.store.report(self.report["id"])
        self.assertEqual((row["name"], row["date_column"], row["date_source"]),
                         ("Продажи 2", "Дата операции", "manual"))
        self.store.save_report(self.report["id"], "Продажи 2", "")
        self.assertEqual(self.store.report(self.report["id"])["date_source"], "auto")

    def test_used_items_cannot_be_deleted(self):
        run, _ = self.compare(BASE, BASE)
        self.assertRaises(ValueError, self.store.delete_report, self.report["id"])
        self.assertRaises(ValueError, self.store.delete_system, self.report["system_id"])
        self.store.delete_run(run["id"])
        for r in self.store.reports(self.report["system_id"]):
            self.store.delete_report(r["id"])
        self.store.delete_system(self.report["system_id"])
        self.assertEqual([s["name"] for s in self.store.systems()], ["Биллинг"])

    # ---------- сверка ----------

    def test_run_with_month_breakdown(self):
        run, status = self.compare(BASE, changed())
        self.assertEqual((status, run["by_month"], run["date_column"], run["unplaced"]),
                         ("diff", 1, "Дата", 0))
        self.assertEqual((run["rows_a"], run["rows_b"], run["exact"], run["changed"], run["only_a"]),
                         (6, 5, 3, 2, 1))
        self.assertEqual(self.months(run), {"2026-07": (1, 0), "2026-08": (0, 0), "2026-09": (2, 0)})
        self.assertEqual((run["side_a"], run["side_b"], run["title"]), ("dk", "ng", "Продажи, CRM"))
        self.assertEqual(self.store.report(self.report["id"])["date_column"], "Дата")
        self.assertEqual(service.run_label(run), ("bad", "3 расхождения"))
        self.assertEqual(service.run_title(run), "Продажи · CRM")
        self.assertEqual(service.period_text(run), "01.07.2026 — 30.09.2026")
        self.assertEqual(service.report_file_name(run), "Продажи_CRM_2026-07-01_2026-09-30.xlsx")

        groups, total, pages = service.load_page(self.store, run)
        rows = groups[0]["rows"]
        self.assertEqual((total, pages), (3, 1))
        self.assertEqual([(r["kind"], r["month"]) for r in rows],
                         [("changed", "2026-07"), ("changed", "2026-09"), ("only_a", "2026-09")])
        first = rows[0]
        self.assertEqual((first["row_a"], first["row_b"], first["count"], first["fields"]),
                         (2, 2, "1 отличие", "4 поля"))
        self.assertEqual(first["diffs"], [(("C", "Сумма"), "10", "11", "значения разные")])
        self.assertEqual(first["ident"][0], (("A", "ID"), "a1"))       # по чему узнать запись
        self.assertEqual([f[3] for f in first["full"]], [False, False, True, False])
        self.assertEqual(rows[2]["a"][0], "a6")
        self.assertIsNone(rows[2]["b"])

        july, total, _ = service.load_page(self.store, run, "2026-07")
        self.assertEqual((total, len(july[0]["rows"])), (1, 1))
        self.assertEqual(service.load_page(self.store, run, "2026-08")[:2], ([], 0))

        # Excel-отчёт сохранён в папке данных и относится к этой сверке
        book = load_workbook(self.store.report_path(run["id"]))
        self.assertEqual(book.sheetnames, ["Итог"])
        self.assertIn("Продажи, CRM: dk ↔ ng", book["Итог"]["A1"].value)
        self.assertTrue(Path(run["file_a"]).name == "выгрузка_dk.xlsx")

    def test_detailed_sheets_on_request(self):
        run, _ = self.compare(BASE, changed(), full=True)
        book = load_workbook(self.store.report_path(run["id"]))
        self.assertEqual(book.sheetnames, ["Итог", "Все проверки", "Описание различий",
                                           "Различия по ячейкам", "Столбцы"])
        self.assertEqual(book["Различия по ячейкам"].max_row, 4)       # заголовок и три расхождения

    def test_empty_cell_is_marked(self):
        other = [list(r) for r in BASE]
        other[1][3] = None
        run, _ = self.compare(BASE, other)
        row = service.load_page(self.store, run)[0][0]["rows"][0]
        self.assertEqual(row["diffs"], [(("D", "Клиент"), "Петров", None, "не заполнено в ng")])

    def test_run_without_differences(self):
        run, status = self.compare(BASE, list(reversed(BASE)))
        self.assertEqual((status, run["by_month"], run["exact"]), ("ok", 1, 6))
        self.assertEqual(service.run_label(run), ("ok", "совпало"))
        self.assertEqual(self.months(run), {"2026-07": (0, 0), "2026-08": (0, 0), "2026-09": (0, 0)})

    def test_run_without_date_column(self):
        head = ["ID", "Сумма"]
        run, status = self.compare([["a1", 1], ["a2", 2]], [["a1", 1], ["a2", 3]], head=head)
        self.assertEqual((status, run["by_month"], run["unplaced"]), ("diff", 0, 1))
        self.assertEqual(self.months(run), {"2026-07": (0, 0), "2026-08": (0, 0), "2026-09": (0, 0)})

    def test_manual_date_column(self):
        head = ["ID", "Дата", "Сумма", "Дата оплаты"]
        base = [["a1", datetime(2026, 7, 5), 10, datetime(2026, 8, 1)],
                ["a2", datetime(2026, 7, 6), 20, datetime(2026, 9, 2)]]
        other = [list(r) for r in base]
        other[0][2] = 11
        self.store.save_report(self.report["id"], "Продажи", "дата оплаты")
        run, _ = self.compare(base, other, head=head)
        self.assertEqual({k: v[0] for k, v in self.months(run).items()},
                         {"2026-07": 0, "2026-08": 1, "2026-09": 0})
        row = self.store.report(self.report["id"])
        self.assertEqual((row["date_column"], row["date_source"]), ("дата оплаты", "manual"))

    def test_partial_month_and_rows_outside_period(self):
        run, _ = self.compare(BASE, changed(), period=(date(2026, 7, 10), date(2026, 8, 31)))
        self.assertEqual(self.months(run), {"2026-07": (0, 1), "2026-08": (0, 0)})
        self.assertEqual(run["unplaced"], 3)            # 5 июля и сентябрь - вне периода
        self.assertEqual(service.load_page(self.store, run, "none")[1], 3)

    def test_period_is_optional_and_taken_from_file(self):
        run, status = self.compare(BASE, changed(), period=(None, None))
        self.assertEqual((run["period_from"], run["period_to"], run["period_auto"]),
                         ("2026-07-01", "2026-09-30", 1))
        self.assertEqual((status, run["by_month"], run["unplaced"]), ("diff", 1, 0))
        self.assertEqual(self.months(run), {"2026-07": (1, 0), "2026-08": (0, 0), "2026-09": (2, 0)})

    def test_no_period_and_no_date_column(self):
        head = ["ID", "Сумма"]
        run, status = self.compare([["a1", 1], ["a2", 2]], [["a1", 1], ["a2", 3]],
                                   period=(None, None), head=head)
        self.assertEqual((status, run["period_from"], run["period_auto"], run["by_month"]),
                         ("diff", "", 0, 0))
        self.assertEqual(self.months(run), {})
        self.assertEqual(service.period_text(run), "не указан")
        self.assertTrue(self.store.report_path(run["id"]).is_file())

    def test_period_from_file_ignores_wild_dates(self):
        wild = [list(r) for r in BASE]
        wild[0][1] = datetime(1990, 1, 1)
        run, status = self.compare(wild, wild, period=(None, None))
        self.assertEqual((status, run["period_from"], run["by_month"]), ("ok", "", 0))
        note = json.loads(run["summary"])["period_note"]
        self.assertIn("в столбце «Дата» даты с 01.01.1990 по 28.09.2026", note)
        self.assertIn("больше допустимых 60", note)

    def test_options_are_applied(self):
        numbered = [[k + 1, *r] for k, r in enumerate(BASE)]
        shuffled = [[k + 1, *r] for k, r in enumerate(reversed(BASE))]
        head = ["№", *HEAD]
        self.assertEqual(self.compare(numbered, shuffled, head=head)[1], "diff")
        self.assertEqual(self.compare(numbered, shuffled, head=head, skip_first_column=True)[1], "ok")
        self.assertEqual(self.compare(numbered, shuffled, head=head, ignore=("№",))[1], "ok")
        run, _ = self.compare(BASE, changed(), keys=["ID"])
        self.assertEqual((run["changed"], run["only_a"]), (2, 1))

    def test_run_without_report(self):
        """Сверка без справочника: названия берутся из имён файлов, как в скрипте."""
        run, status = self.compare(BASE, changed(), report=None, period=(None, None),
                                   names=("продажи_dk_20261113.xlsx", "продажи_ng_20261113.xlsx"))
        self.assertEqual((status, run["report"], run["system"], run["title"]),
                         ("diff", None, None, "продажи"))
        self.assertEqual((run["side_a"], run["side_b"], run["by_month"]), ("dk", "ng", 1))
        self.assertEqual(service.run_title(run), "продажи")
        self.assertEqual(service.report_file_name(run), "продажи_differences.xlsx")
        run, _ = self.compare(BASE, BASE, report=None, names=("март_итог.xlsx", "апрель_свод.xlsx"))
        self.assertEqual((run["side_a"], run["in_b"], service.run_title(run)),
                         ("Файл 1", "файле 2", "Сверка двух файлов"))

    def test_broken_file_is_saved_as_error(self):
        good, bad = self.xlsx(BASE), self.folder / "broken.xlsx"
        bad.write_text("это не Excel", encoding="utf-8")
        report = self.store.report(self.report["id"])
        run_id = self.store.create_run(report, good, bad, None, None, "{}")
        self.assertEqual(self.store.run(run_id)["status"], "error")     # пока не завершена
        status = service.execute(self.store, run_id, good, bad, report, None, None, ce.Options())
        run = self.store.run(run_id)
        self.assertEqual((status, run["status"]), ("error", "error"))
        self.assertIn("Это не файл Excel", run["error_text"])
        self.assertEqual(service.run_label(run), ("err", "ошибка"))
        self.assertFalse(self.store.report_path(run_id).exists())
        self.assertEqual(service.load_page(self.store, run)[:2], ([], 0))

    # ---------- история ----------

    def test_history_and_delete(self):
        first, _ = self.compare(BASE, changed())
        second, _ = self.compare(BASE, BASE, report=None)
        rows, total = self.store.runs()
        self.assertEqual((total, [r["id"] for r in rows]), (2, [second["id"], first["id"]]))
        self.assertEqual(self.store.runs(status="ok")[1], 1)
        self.assertEqual(self.store.runs(status="diff")[1], 1)
        self.assertEqual(self.store.runs(report_id=self.report["id"])[1], 1)
        self.assertEqual(self.store.runs(system_id=self.report["system_id"])[1], 1)
        self.assertEqual(self.store.runs(system_id=self.report["system_id"], status="ok")[1], 0)
        self.store.delete_run(first["id"])
        self.assertEqual(self.store.runs()[1], 1)
        self.assertEqual(self.store.db.execute("SELECT COUNT(*) FROM run_diffs").fetchone()[0], 0)
        self.assertEqual(self.store.months(first["id"]), [])
        self.assertFalse(self.store.report_path(first["id"]).exists())

    def test_pages(self):
        rows = [[f"id{k}", datetime(2026, 7, 1 + k % 28), k, "x"] for k in range(120)]
        other = [[r[0], r[1], r[2] + 1, r[3]] for r in rows]
        run, _ = self.compare(rows, other, keys=["ID"])
        groups, total, pages = service.load_page(self.store, run, page=3)
        self.assertEqual((total, pages, len(groups[0]["rows"])), (120, 3, 20))

    def test_history_survives_reopening(self):
        run, _ = self.compare(BASE, changed())
        self.store.close()
        self.store = Store(self.folder / "data")
        again = self.store.run(run["id"])
        self.assertEqual((again["status"], again["changed"]), ("diff", 2))
        self.assertEqual(len(service.load_page(self.store, again)[0][0]["rows"]), 3)


if __name__ == "__main__":
    unittest.main(verbosity=1)
