#!/usr/bin/env python3
"""
Проверка, что pandas_lite ведёт себя так же, как настоящий pandas.

Один и тот же compare_excel.py загружается дважды - с настоящим pandas (чтение
через openpyxl) и с pandas_lite - и запускается на одних и тех же файлах.
Сравнивается:
    1. чтение книг: названия столбцов и каждое значение вместе с типом;
    2. разбор дат (to_datetime);
    3. готовые отчёты: каждая ячейка, её тип, формат, шрифт, заливка, рамки,
       объединения, ширины столбцов, высоты строк, фильтры, закрепление.

Нужен настоящий pandas (только для этой проверки):
    pip install pandas openpyxl python-dateutil
    python tests/compare_with_pandas.py
"""

import contextlib
import datetime as dt
import importlib.util
import io
import random
import re
import sys
import tempfile
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import pandas as real                      # noqa: E402
import pandas_lite as lite                 # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402

warnings.simplefilter("ignore")
failures = []


def fail(text):
    failures.append(text)
    if len(failures) <= 8:
        print("  ОТЛИЧИЕ:", text[:500])


def load_engine(pandas_module, name):
    """compare_excel.py с указанным модулем в роли pandas (calamine отключена)."""
    saved = {k: sys.modules.get(k) for k in ("pandas", "python_calamine")}
    sys.modules["pandas"], sys.modules["python_calamine"] = pandas_module, None
    try:
        spec = importlib.util.spec_from_file_location(name, ROOT / "compare_excel.py")
        module = importlib.util.module_from_spec(spec)
        sys.modules[name] = module
        spec.loader.exec_module(module)
    finally:
        for k, v in saved.items():
            if v is None:
                sys.modules.pop(k, None)
            else:
                sys.modules[k] = v
    return module


def plain(v):
    """Значение с типом, пригодное для сравнения (NaN равен NaN, numpy -> Python)."""
    if hasattr(v, "item") and type(v).__module__ == "numpy":
        v = v.item()
    if v is real.NaT or v is lite.NaT:
        return ("NaT", None)
    if isinstance(v, real.Timestamp):
        v = v.to_pydatetime()
    if isinstance(v, float) and v != v:
        return ("float", "nan")
    return (type(v).__name__, v)


# ---------- 1. чтение ----------

WORDS = ["Иванов", "ООО Ромашка", " пробел в начале", "пробел в конце ", "NA", "NULL", "None",
         "n/a", "", " ", "0810", "007", "2025", "1.50", "1,5", "=A1+1", "True", "да",
         "13.11.2026", "2026-11-13 10:00:00", "Ёж\tтаб", "строка\nв две", "'апостроф", "-"]


def random_value(rnd):
    kind = rnd.randrange(14)
    if kind == 0:
        return None
    if kind == 1:
        return rnd.randrange(-5, 2000)
    if kind == 2:
        return float(rnd.randrange(0, 50))            # целое, записанное дробным
    if kind == 3:
        return round(rnd.uniform(-1000, 1000), rnd.randrange(0, 6))
    if kind == 4:
        return rnd.choice([0.1 + 0.2, 1e-9, 1e15, 123456789012345678, 0.30000000000000004])
    if kind == 5:
        return rnd.choice([True, False])
    if kind == 6:
        return dt.datetime(2026, rnd.randrange(1, 13), rnd.randrange(1, 28),
                           rnd.randrange(24), rnd.randrange(60), rnd.randrange(60))
    if kind == 7:
        return dt.date(2025, rnd.randrange(1, 13), rnd.randrange(1, 28))
    if kind == 8:
        return dt.time(rnd.randrange(24), rnd.randrange(60))
    if kind == 9:
        return dt.timedelta(hours=rnd.randrange(1, 60), minutes=rnd.randrange(60))
    if kind == 10:
        return rnd.choice(["#N/A", "#DIV/0!", "#VALUE!"])    # ячейка с ошибкой
    return rnd.choice(WORDS)


def random_header(rnd, width):
    style = rnd.randrange(8)
    if style == 0:
        return [f"Столбец {k}" for k in range(width)]
    if style == 1:
        return [rnd.choice(["ID", "Сумма", "Дата", None, "", " Имя ", "Имя", 2025]) for _ in range(width)]
    if style == 2:
        return [2020 + k for k in range(width)]                  # только целые
    if style == 3:
        return [2020 + k if k % 2 else 1.5 + k for k in range(width)]   # целые и дробные
    if style == 4:
        return [dt.datetime(2026, 1, 1 + k) for k in range(width)]
    if style == 5:
        return [None] * width
    if style == 6:
        return ["x"] * width                                     # одни повторы
    return [random_value(rnd) for _ in range(width)]


def make_random_book(path, rnd):
    book = Workbook()
    book.remove(book.active)
    for s in range(rnd.randrange(1, 4)):
        ws = book.create_sheet(f"Лист {s}" if rnd.random() < 0.8 else f"Sheet{s}")
        shape = rnd.randrange(10)
        if shape == 0:
            continue                                             # пустой лист
        width, height = rnd.randrange(1, 9), rnd.randrange(0, 12)
        for _ in range(rnd.randrange(0, 3)):                     # строки перед заголовком
            ws.append([random_value(rnd) for _ in range(rnd.randrange(0, width + 2))])
        ws.append(random_header(rnd, width))
        for _ in range(height):
            ws.append([random_value(rnd) for _ in range(rnd.randrange(0, width + 3))])
        if shape == 1:                                           # хвост из пустых ячеек
            ws.cell(ws.max_row + 3, width + 4).value = None
            ws.cell(ws.max_row + 1, 1).value = "  "
    if not book.worksheets:
        book.create_sheet("Пусто")
    book.save(path)


def check_reading(folder, count=300):
    print(f"1. Чтение книг: {count} случайных файлов, строки заголовков 1-3")
    rnd = random.Random(20261113)
    cells = 0
    for n in range(count):
        path = folder / f"read_{n}.xlsx"
        make_random_book(path, rnd)
        for header in (0, 1, 2):
            kw = dict(sheet_name=None, header=header, dtype=object,
                      keep_default_na=False, na_filter=False)
            results = []
            for module in (real, lite):
                try:
                    results.append(module.read_excel(path, **kw))
                except Exception as e:
                    results.append(type(e).__name__)
            a, b = results
            where = f"{path.name}, header={header}"
            if isinstance(a, str) or isinstance(b, str):
                if not (isinstance(a, str) and a == b):
                    fail(f"{where}: pandas {a if isinstance(a, str) else 'прочитал'}, "
                         f"замена {b if isinstance(b, str) else 'прочитала'}")
                continue
            if list(a) != list(b):
                fail(f"{where}: листы {list(a)} / {list(b)}")
                continue
            for sheet in a:
                da, db = a[sheet], b[sheet]
                ca, cb = [plain(c) for c in da.columns], [plain(c) for c in db.columns]
                if ca != cb or len(da) != len(db):
                    fail(f"{where}, {sheet}: столбцы {ca} / {cb}, строк {len(da)} / {len(db)}")
                    continue
                if [str(c) for c in da.columns] != [str(c) for c in db.columns]:
                    fail(f"{where}, {sheet}: названия столбцов выглядят по-разному")
                for k in range(len(ca)):
                    va = [plain(v) for v in da.iloc[:, k].tolist()]
                    vb = [plain(v) for v in db._data[k]]
                    cells += len(va)
                    if va != vb:
                        fail(f"{where}, {sheet}, столбец {k}: {va[:6]} / {vb[:6]}")
    print(f"   сравнено значений: {cells}")


# ---------- 2. даты ----------

DATE_TEXTS = """
2026-11-13 10:00:00|2026-11-13 10:07:30|2026-11-13T10:00:00|2026-11-13|2026-11-13 10:00
2026-11-13 10:00:00.123|2026-11-13 10:00:00.123456|2026-11-13T10:00:00Z|2026-11-13T10:00:00+03:00
2026-11-13 10:00:00+0300|20261113|20261113100000|20261113T100000|2026-11|2026|2025|0810|1999
13.11.2026|13.11.2026 10:07:30|13.11.2026 10:07|05.11.2026|05.11.2026 10:00:00|1.2.2026|31.12.2026
11/13/2026|13/11/2026|13/11/2026 10:00|05/11/2026 08:15:00|12-31-2026|31-12-2026|2026/11/13
2026/11/13 10:00:00|2026.11.13|13-Nov-26|13 Nov 2026|13 Nov 2026 10:00|Nov 13, 2026|November 13, 2026
Fri Nov 13 10:00:00 2026|13 November 2026 10:00:00|Nov 2026|11/2026|11-2026|2026Q1|2026Q4|2026-Q2|1Q2026
10:00|10:00:05|9:05|23:59:59|dk|ng|10|11|1.5|-5|+5|1e5|999|1000|1001|12345|100000|v2|v3
да|нет|март|Пт 13|a|T|am|may|None|NULL|n/a|-|abc 123|ООО Ромашка|2026-13-45|32.01.2026|00.00.0000
2026-11-13 25:00:00|13.11.26|13.11.26 10:00|1,5|5 5|1 2 3|202611|2026 11 13|13 11 2026|0|00|007|01.02
 2026-11-13 10:00:00|2026-11-13 10:00:00 |13.11.2026 |Период: 2026|2026 год|Q1|v1.2.3|1.2.3|10.11
"""


def check_dates():
    texts = [t for line in DATE_TEXTS.strip("\n").split("\n") for t in line.split("|")]
    values = texts + [dt.datetime(2026, 11, 13, 10), dt.datetime(2026, 11, 13, 10, 7, 30),
                      dt.date(2026, 1, 1), dt.time(10, 0), dt.timedelta(hours=5), True, 5, 1.5]
    print(f"2. Разбор дат: {len(values)} значений, все пары между собой ({len(values) ** 2})")
    engines = [load_engine(real, "engine_real_dates"), load_engine(lite, "engine_lite_dates")]
    wrong = {}
    for v1 in values:
        for v2 in values:
            a, b = (e.tech_comment(v1, v2) for e in engines)
            if a != b:
                wrong.setdefault(repr(v2 if repr(v1) in wrong else v1), []).append((v1, v2, a, b))
    for key, cases in sorted(wrong.items(), key=lambda x: -len(x[1])):
        v1, v2, a, b = cases[0]
        fail(f"пояснение, значение {key}: расходится в {len(cases)} парах, например с "
             f"{v2 if repr(v1) == key else v1!r}:\n      pandas: {a}\n      замена: {b}")
    for v in values:      # и сам разбор: дата получилась или нет
        outcome = []
        for module in (real, lite):
            try:
                t = module.to_datetime(v, errors="raise")
                outcome.append("дата" if t is not real.NaT else "NaT")
            except Exception:
                outcome.append("не дата")
        if outcome[0] != outcome[1] and not isinstance(v, (int, float)):
            fail(f"to_datetime({v!r}): pandas - {outcome[0]}, замена - {outcome[1]}")


# ---------- 3. отчёты ----------

def side_styles(border):
    return tuple(getattr(border, s).style for s in ("left", "right", "top", "bottom"))


def book_snapshot(path):
    """Всё содержимое и оформление книги в виде, удобном для сравнения."""
    stamp = re.compile(r"\d{4}-\d\d-\d\d \d\d:\d\d:\d\d")
    book = load_workbook(path)
    result = {"sheets": book.sheetnames}
    for ws in book.worksheets:
        cells = {}
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and ("дата сверки" in v or stamp.fullmatch(v)):
                    v = stamp.sub("<дата>", v)
                cells[c.coordinate] = (
                    plain(v), c.number_format, c.font.b, c.font.i, c.font.sz,
                    c.font.color.rgb if c.font.color else None, c.fill.fgColor.rgb,
                    c.fill.fill_type, side_styles(c.border), c.alignment.horizontal,
                    c.alignment.vertical, c.alignment.wrap_text)
        rules = [(str(rng.sqref), [(r.type, tuple(r.formula), r.dxf.fill.bgColor.rgb)
                                   for r in rule_list])
                 for rng, rule_list in ws.conditional_formatting._cf_rules.items()]
        result[ws.title] = {
            "cells": cells, "dimensions": ws.dimensions,
            "merged": sorted(str(r) for r in ws.merged_cells.ranges),
            "widths": {k: d.width for k, d in ws.column_dimensions.items()},
            "heights": {k: d.height for k, d in ws.row_dimensions.items()},
            "freeze": ws.freeze_panes, "filter": ws.auto_filter.ref, "rules": rules,
            "grid": ws.sheet_view.showGridLines, "selected": ws.sheet_view.tabSelected}
    return result


def diff_snapshots(name, a, b):
    if a["sheets"] != b["sheets"]:
        return fail(f"{name}: листы {a['sheets']} / {b['sheets']}")
    for sheet in a["sheets"]:
        sa, sb = a[sheet], b[sheet]
        for key in sa:
            if key != "cells" and sa[key] != sb[key]:
                fail(f"{name}, лист «{sheet}»: {key}: {sa[key]} / {sb[key]}")
        for coordinate in sorted(set(sa["cells"]) | set(sb["cells"])):
            ca, cb = sa["cells"].get(coordinate), sb["cells"].get(coordinate)
            if ca != cb:
                fail(f"{name}, лист «{sheet}», {coordinate}:\n      pandas: {ca}\n      замена: {cb}")
    return sum(len(a[s]["cells"]) for s in a["sheets"])


def base_rows(n, rnd):
    return [{"ID": f"A{i:05d}", "Клиент": f"ООО Ромашка {i}", "Счёт": f"0{4070 + i}",
             "Сумма": round(100.5 * i, 2), "Кол-во": i % 7, "Доля": rnd.random(),
             "Дата": dt.datetime(2026, 1 + i % 12, 1 + i % 27, i % 24),
             "Активен": i % 3 == 0, "Комментарий": rnd.choice(["", "ок", " проверить ", "NA"]),
             "Время": dt.time(i % 24, i % 60), "Длительность": dt.timedelta(minutes=90 + i),
             "Ошибка": "#N/A" if i % 9 == 0 else i,
             "Сформировано": "13.11.2026 10:00:00"} for i in range(1, n + 1)]


def write_rows(ws, rows, order=None):
    names = order or list(rows[0])
    ws.append(names)
    for r in rows:
        ws.append([r.get(c) for c in names])


def make_pair(folder, name, rnd, n=60, damage=12, stamp_b="13.11.2026 10:07:30", extra=None,
              kinds=range(13)):
    """Два файла: второй - перемешанная копия первого с набором разных повреждений."""
    rows_a = base_rows(n, rnd)
    rows_b = [dict(r, **{"Сформировано": stamp_b}) for r in rows_a]
    for _ in range(damage):
        r = rnd.choice(rows_b)
        kind = rnd.choice(list(kinds))
        if kind == 0:
            r["Сумма"] = r["Сумма"] + 1
        elif kind == 1:
            r["Клиент"] = None
        elif kind == 2:
            r["Клиент"] = str(r["Клиент"]).upper()
        elif kind == 3:
            r["Счёт"] = int(r["Счёт"]) if isinstance(r["Счёт"], str) else r["Счёт"]
        elif kind == 4:
            if isinstance(r["Дата"], dt.datetime):
                r["Дата"] = r["Дата"] + dt.timedelta(days=1)
        elif kind == 5:
            r["Активен"] = not r["Активен"]
        elif kind == 6:
            r["Комментарий"] = "«новый»  текст"
        elif kind == 7:
            r["Доля"] = None
        elif kind == 8:
            if isinstance(r["Дата"], dt.datetime):
                r["Дата"] = r["Дата"].strftime("%d.%m.%Y %H:%M:%S")
        elif kind == 9:
            r["Время"] = dt.time(23, 59, 30)
        elif kind == 10:
            r["Длительность"] = None if r["Длительность"] is None else r["Длительность"] * 2
        elif kind == 11:
            r["Кол-во"] = r["Кол-во"] + 1
        else:
            r["Дата"] = None
    if len(set(kinds)) > 4:        # в узких наборах строки не пропадают: отличия одного типа
        removed = [rows_b.pop(rnd.randrange(len(rows_b))) for _ in range(min(3, damage))]
        rows_b += [dict(removed[0], ID=f"B{k}", Клиент=f"ИП Новый {k}")
                   for k in range(min(2, damage))]
    rnd.shuffle(rows_b)
    order_b = list(rows_a[0])
    rnd.shuffle(order_b)
    paths = [folder / f"{name}_dk_20261113.xlsx", folder / f"{name}_ng_20261113.xlsx"]
    for path, rows, order in ((paths[0], rows_a, None), (paths[1], rows_b, order_b)):
        book = Workbook()
        write_rows(book.active, rows, order)
        if extra:
            extra(book, path is paths[0])
        book.save(path)
    return paths


def extra_sheets(book, first):
    ws = book.create_sheet("Продажи")
    ws.append(["Регион", "План", "Факт", None, "Регион"])
    for k in range(15):
        ws.append([f"Регион {k}", 1000 + k, 900 + k * (1 if first else 2), k, "x"])
    book.create_sheet("Только в первом" if first else "Только во втором").append(["a", "b"])
    ws = book.create_sheet("Числа")
    ws.append(["Код", "Значение", "Только здесь" if first else "А тут другое"])
    for k in range(8):
        ws.append([k, k * 1.5 if first or k != 3 else None, "#N/A" if k == 5 else k])


def options_sets(engine):
    make = engine.Options
    return {
        "обычная": make(),
        "full": make(full=True),
        "sum": make(full=True, sum_cols=engine.parse_columns(["D:F"])),
        "key": make(keys=["ID"], full=True),
        "by_order": make(by_order=True),
        "text_numbers": make(text_numbers=True, ignore_case=True, tolerance=0.5, full=True),
        "skip": make(skip_first_column=True, ignore=("Комментарий",), similarity=0.8, full=True),
    }


def check_reports(folder):
    print("3. Отчёты: один и тот же compare_excel.py с pandas и с заменой")
    rnd = random.Random(7)
    pairs = {
        "report01": make_pair(folder, "report01", rnd),
        "report02": make_pair(folder, "report02", rnd, n=25, damage=0, stamp_b="13.11.2026 10:00:00"),
        "report03": make_pair(folder, "report03", rnd, n=40, damage=6, extra=extra_sheets),
        "report04": make_pair(folder, "report04", rnd, n=200, damage=120, stamp_b="другая система"),
        "report05": make_pair(folder, "report05", rnd, n=3200, damage=9000),   # подробные листы сами
        "report06": make_pair(folder, "report06", rnd, n=50, damage=20, kinds=(0, 7, 11)),   # только числа
        "report07": make_pair(folder, "report07", rnd, n=50, damage=20, kinds=(4, 12)),      # только даты
        "report08": make_pair(folder, "report08", rnd, n=50, damage=20, kinds=(5,)),         # только да/нет
        "report09": make_pair(folder, "report09", rnd, n=50, damage=20, kinds=(9, 10)),      # время
    }
    engines = {"pandas": load_engine(real, "engine_real"), "замена": load_engine(lite, "engine_lite")}
    total = 0
    for pair_name, (f1, f2) in pairs.items():
        names = ["обычная"] if pair_name == "report05" else list(options_sets(engines["pandas"]))
        for opt_name in names:
            if opt_name == "sum" and pair_name == "report03":
                continue          # на дополнительных листах нет столбцов D:F
            shots, logs = [], []
            for label, engine in engines.items():
                o = options_sets(engine)[opt_name]
                out = folder / f"{pair_name}_{opt_name}_{'a' if label == 'pandas' else 'b'}.xlsx"
                text = io.StringIO()
                try:
                    with contextlib.redirect_stdout(text):
                        engine.run_pair(f1, f2, o, None, out)
                    shots.append(book_snapshot(out))
                except Exception as e:
                    shots.append(f"{type(e).__name__}: {e}")
                logs.append([s for s in text.getvalue().split("\n")
                             if not s.startswith(("Отчёт: ", "  Чтение файлов можно ускорить"))])
            name = f"{pair_name}/{opt_name}"
            if isinstance(shots[0], str) or isinstance(shots[1], str):
                if shots[0] != shots[1]:
                    fail(f"{name}: pandas - {shots[0] if isinstance(shots[0], str) else 'отчёт'}; "
                         f"замена - {shots[1] if isinstance(shots[1], str) else 'отчёт'}")
                continue
            if logs[0] != logs[1]:
                fail(f"{name}: текст в консоли отличается")
            cells = diff_snapshots(name, *shots)
            total += cells or 0
            print(f"   {name}: листов {len(shots[0]['sheets'])}, ячеек {cells}")
    print(f"   всего сравнено ячеек отчётов: {total}")


def main():
    print(f"pandas {real.__version__}, Python {sys.version.split()[0]}")
    with tempfile.TemporaryDirectory() as tmp:
        folder = Path(tmp)
        check_reading(folder)
        check_dates()
        check_reports(folder)
    print(f"\nИтог: отличий - {len(failures)}" if failures else "\nИтог: отличий нет")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
