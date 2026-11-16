"""
Запуск сверки с записью в историю и подготовка результата к показу в окне.

Сама сверка и Excel-отчёт - это compare_excel.py (compare_files и write_report).
Здесь только то, что вокруг: период сверки, столбец с датой, раскладка
расхождений по месяцам, карточки расхождений.
"""

import calendar
import json
import re
from datetime import date, datetime, time, timedelta
from pathlib import Path

from engine import ce

MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август",
          "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
MAX_AUTO_MONTHS = 60    # период из файла длиннее этого считаем ошибкой в датах
PAGE = 50               # расхождений на одной странице окна
DATE_FORMATS = ((10, "%d.%m.%Y"), (10, "%Y-%m-%d"), (10, "%d/%m/%Y"), (7, "%Y-%m"), (7, "%m.%Y"))


# ---------- даты и месяцы ----------

def to_date(v):
    """Дата из ячейки (в том числе записанная текстом) или None."""
    if ce.is_empty(v):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        s = v.strip()
        for length, fmt in DATE_FORMATS:
            # после даты может идти время: «13.11.2026 10:00» или «2026-11-13T10:00»
            if len(s) == length or (length == 10 and len(s) > 10 and s[10] in " T"):
                try:
                    return datetime.strptime(s[:length], fmt).date()
                except ValueError:
                    pass
    return None


def parse_user_date(text):
    """Дата, введённая в окне: ДД.ММ.ГГГГ или ГГГГ-ММ-ДД. Пустая строка - None."""
    text = (text or "").strip()
    if not text:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Не понял дату «{text}». Пишите так: 31.12.2026")


def quick_period(kind, today=None):
    """Готовый период: month - прошлый месяц, quarter - прошлый квартал, year - с начала года."""
    today = today or date.today()
    if kind == "year":
        return date(today.year, 1, 1), today
    months = 1 if kind == "month" else 3
    first_now = (today.month - 1) // months * months      # номер первого месяца периода (с нуля)
    last = date(today.year, first_now + 1, 1) - timedelta(days=1)
    index = last.year * 12 + last.month - months
    return date(index // 12, index % 12 + 1, 1), last


def month_key(d):
    return f"{d.year:04d}-{d.month:02d}"


def month_bounds(key):
    year, month = int(key[:4]), int(key[5:7])
    return date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1])


def month_title(key):
    return f"{MONTHS[int(key[5:7]) - 1]} {key[:4]}"


def shift_month(key, delta):
    index = int(key[:4]) * 12 + int(key[5:7]) - 1 + delta
    return f"{index // 12:04d}-{index % 12 + 1:02d}"


def months_between(first, last):
    result, key, end = [], month_key(first), month_key(last)
    while key <= end:
        result.append(key)
        key = shift_month(key, 1)
    return result


# ---------- столбец с датой ----------

def find_date_column(df, period_from, period_to):
    """
    Ищет в листе столбец с датами записей: почти все значения - даты, они не
    одинаковые, и не меньше половины попадает в период сверки. Если период не
    задан, подходит первый такой столбец. Возвращает название столбца или None.
    """
    best, best_score = None, 0.0
    for column in df.columns:
        values = [v for v in df[column].tolist()[:2000] if not ce.is_empty(v)]
        if not values:
            continue
        dates = [d for d in map(to_date, values) if d is not None]
        if len(dates) < 0.8 * len(values):
            continue
        if len(values) > 1 and len({str(v) for v in values}) == 1:
            continue          # одно значение на весь файл: это не дата записи
        if period_from is None:
            return column
        score = sum(period_from <= d <= period_to for d in dates) / len(dates)
        if score >= 0.5 and score > best_score:
            best, best_score = column, score
    return best


def resolve_column(df, name):
    """Столбец листа по названию из справочника (регистр и пробелы по краям не важны)."""
    wanted = str(name).strip().casefold()
    for column in df.columns:
        if str(column).strip().casefold() == wanted:
            return column
    return None


# ---------- значения для показа ----------

def show(v):
    """Значение ячейки текстом - так, как его привыкли видеть в Excel."""
    if ce.is_empty(v):
        return None
    if isinstance(v, datetime):
        plain = (v.hour, v.minute, v.second) == (0, 0, 0)
        return v.strftime("%d.%m.%Y" if plain else "%d.%m.%Y %H:%M:%S")
    if isinstance(v, date):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, bool):
        return "ИСТИНА" if v else "ЛОЖЬ"
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else f"{v:.12g}"
    return str(v)


def dump(value):
    return json.dumps(value, ensure_ascii=False, default=show)


def number(n):
    return f"{n or 0:,}".replace(",", " ")


def plural_diff(n):
    return f"{n} {ce.plural(n, 'расхождение', 'расхождения', 'расхождений')}"


# ---------- сверка ----------

def read_and_compare(file_a, file_b, o, nm):
    """compare_files и прочитанные им книги (нужны, чтобы найти столбец с датой)."""
    books, original = [], ce.read_book

    def reading(path, header):
        book = original(path, header)
        books.append(book)
        return book

    ce.read_book = reading
    try:
        res = ce.compare_files(file_a, file_b, o, nm)
    finally:
        ce.read_book = original
    return res, books[-2], books[-1]


def execute(store, run_id, file_a, file_b, report, period_from, period_to, o, names=None):
    """
    Сверяет два файла и записывает результат в историю. Возвращает статус:
    ok - совпало, diff - есть расхождения, error - сверка не выполнена.
    report - строка справочника отчётов или None (сверка без справочника),
    o - параметры сверки (Options), names - свои названия файлов или None.
    """
    try:
        result = _compare(store, run_id, Path(file_a), Path(file_b), report,
                          period_from, period_to, o, names)
    except Exception as e:
        store.db.rollback()
        if isinstance(e, PermissionError):
            text = (f"Нет доступа к файлу {e.filename}. Если он открыт в Excel, закройте "
                    "его и запустите сверку снова")
        elif isinstance(e, (ValueError, OSError)):
            text = str(e) or type(e).__name__
        else:
            text = f"({type(e).__name__}): {e}"
        result = {"status": "error", "error_text": text[:2000]}
    store.finish_run(run_id, result)
    return result["status"]


def _compare(store, run_id, file_a, file_b, report, period_from, period_to, o, names):
    nm = ce.derive_names(file_a, file_b, names)
    if report:
        system = store.db.execute("SELECT name FROM systems WHERE id = ?",
                                  (report["system_id"],)).fetchone()[0]
        nm.prefix = f"{report['name']}, {system}"
    res, book1, book2 = read_and_compare(file_a, file_b, o, nm)
    ce.write_report(store.report_path(run_id), res, file_a, file_b, nm, o)
    first_row = o.header + 2       # номер строки Excel для первой строки данных

    # --- столбец с датой на каждом листе ---
    sheets, found = {}, None
    manual = report["date_column"] if report and report["date_source"] == "manual" else None

    def sheet_info(label):
        """Для листа: по каждому файлу (значения столбца с датой или None)."""
        if label not in sheets:
            if label in book1 or label in book2:
                frames = (book1.get(label), book2.get(label))
            elif len(book1) == 1 and len(book2) == 1:      # один лист с разными названиями
                frames = (next(iter(book1.values())), next(iter(book2.values())))
            else:
                frames = (None, None)
            info = []
            for df in frames:
                column = None
                if df is not None:
                    column = (resolve_column(df, manual) if manual
                              else find_date_column(df, period_from, period_to))
                info.append((column, None if column is None else df[column].tolist()))
            sheets[label] = info
        return sheets[label]

    for tb in res.tables:
        for column, _ in sheet_info(tb["sheet"]):
            if column is not None:
                found = found or str(column).strip()

    # Период не задан: берём его из дат в файлах, целыми месяцами
    period_auto, period_note = False, None
    if period_from is None:
        dates = [d for info in sheets.values() for _, values in info if values is not None
                 for d in map(to_date, values) if d is not None]
        span = len(months_between(min(dates), max(dates))) if dates else 0
        if 0 < span <= MAX_AUTO_MONTHS:
            period_from = month_bounds(month_key(min(dates)))[0]
            period_to = month_bounds(month_key(max(dates)))[1]
            period_auto = True
        else:
            if span:           # даты есть, но разброс слишком велик
                period_note = (f"в столбце «{found}» даты с {min(dates):%d.%m.%Y} по "
                               f"{max(dates):%d.%m.%Y} — это {span} "
                               f"{ce.plural(span, 'месяц', 'месяца', 'месяцев')}, больше "
                               f"допустимых {MAX_AUTO_MONTHS}.")
            found = None       # без разбивки по месяцам

    def month_of(label, row_a, row_b):
        if period_from is None:
            return None
        for (column, values), row in zip(sheet_info(label), (row_a, row_b)):
            if values is None or not isinstance(row, int):
                continue
            index = row - first_row
            if 0 <= index < len(values):
                d = to_date(values[index])
                if d is not None and period_from <= d <= period_to:
                    return month_key(d)
            return None
        return None

    # --- строки с расхождениями по месяцам ---
    per_month, unplaced, seen = {}, 0, set()
    for d in res.details:
        row_a, row_b = d["row_a"], d["row_b"]
        if not isinstance(row_a, int) and not isinstance(row_b, int):
            continue                                       # отличие в структуре
        key = (d["Лист"], row_a, row_b)
        if key in seen:
            continue
        seen.add(key)
        month = month_of(*key)
        if month is None:
            unplaced += 1
        else:
            per_month[month] = per_month.get(month, 0) + 1

    db = store.db
    db.execute("DELETE FROM run_months WHERE run_id = ?", (run_id,))
    db.execute("DELETE FROM run_diffs WHERE run_id = ?", (run_id,))
    for month in months_between(period_from, period_to) if period_from else ():
        first, last = month_bounds(month)
        partial = period_from > first or period_to < last
        db.execute("INSERT INTO run_months VALUES (?, ?, ?, ?)",
                   (run_id, month, per_month.get(month, 0), int(partial)))

    # --- сами расхождения для показа в окне ---
    tables = []
    insert = ("INSERT INTO run_diffs (run_id, sheet, kind, row_a, row_b, month, vals_a, vals_b, "
              "kinds, hint) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")
    for tb in res.tables:
        label = tb["sheet"]
        tables.append({"sheet": label, "columns": tb["columns"], "tech": tb["tech"],
                       "whole": tb.get("whole", 0), "changed": tb["changed"],
                       "only_n": tb["only_n"]})
        for row_a, row_b, vals_a, vals_b, kinds, hint in tb["pairs"]:
            db.execute(insert, (run_id, label, "changed", row_a, row_b,
                                month_of(label, row_a, row_b), dump([show(v) for v in vals_a]),
                                dump([show(v) for v in vals_b]), dump(kinds), hint))
        for side, row, values, hint in tb["only"]:
            in_a = side == nm.a
            row_a, row_b = (row, None) if in_a else (None, row)
            shown = dump([show(v) for v in values])
            db.execute(insert, (run_id, label, "only_a" if in_a else "only_b", row_a, row_b,
                                month_of(label, row_a, row_b), shown if in_a else None,
                                None if in_a else shown, "{}", hint))

    if report and not manual and found and found != report["date_column"]:
        store.set_found_column(report["id"], found)

    t = res.totals
    whole = getattr(res, "whole", [])
    summary = {
        "tables": tables, "period_note": period_note,
        "checks": [c for c in res.checks if c["Результат"] != ce.OK or c["Комментарий"]],
        "whole": [(s, c, show(v1), show(v2)) for s, c, v1, v2 in whole],
        "tech": [(s, c, show(v1), show(v2)) for s, c, v1, v2 in res.tech],
        "structure": [d for d in res.descriptions if d["Категория"] == "Структура"],
    }
    return {
        "status": "diff" if res.has_differences else "ok", "error_text": None,
        "title": nm.prefix, "side_a": nm.a, "side_b": nm.b, "in_a": nm.a_in, "in_b": nm.b_in,
        "rows_a": t["rows_a"], "rows_b": t["rows_b"], "exact": t["exact"],
        "changed": t["changed"], "only_a": t["only_a"], "only_b": t["only_b"],
        "cells": t["cells1"] + t["cells2"] + t["cells3"],
        "structure": t["structure"] + t["const_real"],
        "period_from": period_from.isoformat() if period_from else "",
        "period_to": period_to.isoformat() if period_to else "",
        "period_auto": int(period_auto),
        "by_month": int(found is not None), "date_column": found,
        "unplaced": unplaced, "summary": dump(summary),
    }


# ---------- результат для окна ----------

def run_label(run):
    """Короткий итог сверки: (вид, текст). Вид: ok / bad / err."""
    if run["status"] == "error":
        return "err", "ошибка"
    if run["status"] == "ok":
        return "ok", "совпало"
    rows = run["changed"] + run["only_a"] + run["only_b"]
    return "bad", plural_diff(rows) if rows else "расхождения"


def run_title(run):
    if run["report"]:
        return f"{run['report']} · {run['system']}"
    return run["title"] or "Сверка двух файлов"


def format_day(text):
    try:
        return date.fromisoformat(text).strftime("%d.%m.%Y")
    except (TypeError, ValueError):
        return text or ""


def format_moment(text):
    try:
        return datetime.fromisoformat(text).strftime("%d.%m.%Y %H:%M")
    except (TypeError, ValueError):
        return text or ""


def period_text(run):
    if not run["period_from"]:
        return "не указан"
    return f"{format_day(run['period_from'])} — {format_day(run['period_to'])}"


def report_file_name(run):
    """Имя, под которым предлагается сохранить Excel-отчёт сверки."""
    if run["report"]:
        period = (f"{run['period_from']}_{run['period_to']}" if run["period_from"]
                  else run["created_at"][:10])
        name = f"{run['report']}_{run['system']}_{period}.xlsx"
    else:
        name = f"{run['title']}_differences.xlsx" if run["title"] else "differences.xlsx"
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .") or "differences.xlsx"


def split_label(label):
    """«F: Сумма» -> («F», «Сумма»): буква столбца Excel и название."""
    letter, sep, name = str(label).partition(": ")
    return (letter, name) if sep else ("", str(label))


def card(row, group, in_a, in_b):
    """
    Карточка расхождения: по чему узнать запись (ident), какие поля отличаются
    (diffs: поле, значение 1, значение 2, что не так) и вся строка (full).
    """
    notes = {"empty_a": f"не заполнено в {in_a}", "empty_b": f"не заполнено в {in_b}",
             "value": "значения разные", "form": "разная форма записи"}
    labels = [split_label(c) for c in group["columns"]]
    first_whole, first_tech = group["first_whole"], group["first_tech"]
    a, b, kinds = row["a"], row["b"], row["kinds"]
    base = a if a is not None else b

    def get(values, i):
        return values[i] if values is not None and i < len(values) else None

    differs = sorted(int(k) for k in kinds if int(k) < first_whole)
    diffs = [(labels[i], get(a, i), get(b, i), notes.get(kinds[str(i)], ""))
             for i in differs if i < len(labels)]
    # Запись узнаём по первым заполненным полям, которые в обоих файлах одинаковы
    ident = [(labels[i], base[i]) for i in range(min(first_whole, len(base)))
             if str(i) not in kinds and base[i] is not None][:3]
    full = [(labels[i], get(a, i) if a is not None else get(b, i), get(b, i),
             str(i) in kinds or first_whole <= i < first_tech)
            for i in range(len(labels))]
    n = len(diffs)
    return {"diffs": diffs, "ident": ident, "full": full,
            "count": f"{n} {ce.plural(n, 'отличие', 'отличия', 'отличий')}",
            "fields": f"{len(full)} {ce.plural(len(full), 'поле', 'поля', 'полей')}"}


def load_page(store, run, month="", page=1):
    """
    Одна страница расхождений сверки, сгруппированная по листам.
    Возвращает (группы, всего по условию, страниц). Группа: sheet, columns,
    first_tech, first_whole, rows; строка - расхождение вместе с данными карточки.
    """
    summary = json.loads(run["summary"] or "{}")
    tables = {t["sheet"]: t for t in summary.get("tables", [])}
    rows, total = store.diffs(run["id"], month, (page - 1) * PAGE, PAGE)
    groups = []
    for row in rows:
        if not groups or groups[-1]["sheet"] != row["sheet"]:
            t = tables.get(row["sheet"], {"columns": [], "tech": 0, "whole": 0})
            n = len(t["columns"])
            groups.append({"sheet": row["sheet"], "columns": t["columns"],
                           "first_tech": n - t["tech"],
                           "first_whole": n - t["tech"] - t.get("whole", 0), "rows": []})
        item = {"kind": row["kind"], "row_a": row["row_a"], "row_b": row["row_b"],
                "month": row["month"], "hint": row["hint"],
                "a": json.loads(row["vals_a"]) if row["vals_a"] else None,
                "b": json.loads(row["vals_b"]) if row["vals_b"] else None,
                "kinds": json.loads(row["kinds"] or "{}")}
        item.update(card(item, groups[-1], run["in_a"], run["in_b"]))
        groups[-1]["rows"].append(item)
    return groups, total, max(1, -(-total // PAGE))
