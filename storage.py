"""
Хранилище приложения: справочники и история сверок в одном файле SQLite.

Папка с данными (по умолчанию «SverkaExcel» в папке пользователя):
    sverka.db   справочники, история сверок, расхождения для показа в окне
    reports     готовые Excel-отчёты (run_<номер>.xlsx)
Исходные файлы пользователя не копируются: в истории хранятся только пути к ним.
"""

import os
import sqlite3
from datetime import datetime
from pathlib import Path

SCHEMA = """
CREATE TABLE IF NOT EXISTS systems (
    id   INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS reports (
    id          INTEGER PRIMARY KEY,
    system_id   INTEGER NOT NULL REFERENCES systems(id),   -- у каждой системы свои отчёты
    name        TEXT NOT NULL,
    date_column TEXT,                          -- столбец с датой (для разбивки по месяцам)
    date_source TEXT NOT NULL DEFAULT 'auto',  -- auto: найден программой, manual: указан вручную
    UNIQUE (system_id, name)
);

CREATE TABLE IF NOT EXISTS runs (
    id          INTEGER PRIMARY KEY,
    created_at  TEXT NOT NULL,
    report_id   INTEGER REFERENCES reports(id),            -- пусто: сверка без справочника
    system_id   INTEGER REFERENCES systems(id),
    title       TEXT,                          -- название сверки в заголовке отчёта
    side_a      TEXT,                          -- как названы файлы: «Файл 1» или имя системы
    side_b      TEXT,
    in_a        TEXT,                          -- то же после предлога «в»
    in_b        TEXT,
    period_from TEXT NOT NULL DEFAULT '',      -- ГГГГ-ММ-ДД; пусто, если период неизвестен
    period_to   TEXT NOT NULL DEFAULT '',
    period_auto INTEGER DEFAULT 0,             -- 1: период взят из дат в файле
    status      TEXT NOT NULL,                 -- ok / diff / error
    error_text  TEXT,
    file_a      TEXT,                          -- полные пути к исходным файлам
    file_b      TEXT,
    options     TEXT,                          -- параметры сверки (JSON)
    rows_a      INTEGER DEFAULT 0,
    rows_b      INTEGER DEFAULT 0,
    exact       INTEGER DEFAULT 0,
    changed     INTEGER DEFAULT 0,
    only_a      INTEGER DEFAULT 0,
    only_b      INTEGER DEFAULT 0,
    cells       INTEGER DEFAULT 0,
    structure   INTEGER DEFAULT 0,             -- отличий в структуре файла
    by_month    INTEGER DEFAULT 0,             -- 1: расхождения разложены по месяцам
    date_column TEXT,
    unplaced    INTEGER DEFAULT 0,             -- строк с расхождениями без месяца
    summary     TEXT                           -- проверки, таблицы, структура (JSON)
);

CREATE TABLE IF NOT EXISTS run_months (
    run_id   INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    month    TEXT NOT NULL,                    -- ГГГГ-ММ
    bad_rows INTEGER NOT NULL DEFAULT 0,
    partial  INTEGER NOT NULL DEFAULT 0,       -- 1: период покрывает месяц не целиком
    PRIMARY KEY (run_id, month)
);

CREATE TABLE IF NOT EXISTS run_diffs (
    id     INTEGER PRIMARY KEY,
    run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    sheet  TEXT NOT NULL,
    kind   TEXT NOT NULL,                      -- changed / only_a / only_b
    row_a  INTEGER,
    row_b  INTEGER,
    month  TEXT,
    vals_a TEXT,                               -- значения строки (JSON)
    vals_b TEXT,
    kinds  TEXT,                               -- номер столбца -> вид отличия (JSON)
    hint   TEXT
);

CREATE INDEX IF NOT EXISTS runs_report ON runs(report_id, created_at);
CREATE INDEX IF NOT EXISTS diffs_run ON run_diffs(run_id, month);
"""

RUNS = """
    SELECT r.*, rp.name AS report, s.name AS system
    FROM runs r
    LEFT JOIN reports rp ON rp.id = r.report_id
    LEFT JOIN systems s ON s.id = r.system_id
"""


def default_folder():
    """Папка с данными: переменная SVERKA_DATA или «SverkaExcel» в папке пользователя."""
    return Path(os.environ.get("SVERKA_DATA") or Path.home() / "SverkaExcel")


def clean_lines(text):
    """Названия из многострочного поля: по одному на строку, без пустых и повторов."""
    return list(dict.fromkeys(x.strip() for x in (text or "").splitlines() if x.strip()))


class Store:
    """Одно подключение к базе. В другом потоке создаётся свой Store той же папки."""

    def __init__(self, folder=None):
        self.folder = Path(folder) if folder else default_folder()
        (self.folder / "reports").mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(self.folder / "sverka.db", timeout=10)
        self.db.row_factory = sqlite3.Row
        self.db.execute("PRAGMA foreign_keys = ON")
        self.db.executescript(SCHEMA)
        self.db.commit()

    def close(self):
        self.db.close()

    def report_path(self, run_id):
        return self.folder / "reports" / f"run_{run_id}.xlsx"

    # ---------- справочники ----------

    def systems(self):
        return self.db.execute(
            "SELECT s.*, (SELECT COUNT(*) FROM reports p WHERE p.system_id = s.id) AS used "
            "FROM systems s ORDER BY name COLLATE NOCASE").fetchall()

    def reports(self, system_id=None):
        """Отчёты одной системы (или все) с числом сверок по каждому."""
        where = "WHERE p.system_id = ?" if system_id is not None else ""
        return self.db.execute(
            "SELECT p.*, s.name AS system, "
            "(SELECT COUNT(*) FROM runs r WHERE r.report_id = p.id) AS used "
            f"FROM reports p JOIN systems s ON s.id = p.system_id {where} "
            "ORDER BY s.name COLLATE NOCASE, p.name COLLATE NOCASE",
            () if system_id is None else (system_id,)).fetchall()

    def report(self, report_id):
        return self.db.execute("SELECT * FROM reports WHERE id = ?", (report_id,)).fetchone()

    def _add(self, sql, names, *first):
        added, repeated = 0, []
        for name in names:
            try:
                self.db.execute(sql, (*first, name))
                added += 1
            except sqlite3.IntegrityError:
                repeated.append(name)
        self.db.commit()
        return added, repeated

    def add_systems(self, text):
        """Добавляет системы (каждая с новой строки). Возвращает (добавлено, повторы)."""
        return self._add("INSERT INTO systems (name) VALUES (?)", clean_lines(text))

    def add_reports(self, system_id, text):
        return self._add("INSERT INTO reports (system_id, name) VALUES (?, ?)",
                         clean_lines(text), system_id)

    def rename_system(self, system_id, name):
        name = name.strip()
        if not name:
            raise ValueError("Название не может быть пустым.")
        try:
            self.db.execute("UPDATE systems SET name = ? WHERE id = ?", (name, system_id))
            self.db.commit()
        except sqlite3.IntegrityError:
            raise ValueError(f"«{name}» уже есть в списке.") from None

    def save_report(self, report_id, name, date_column):
        """Название и столбец с датой. Столбец, вписанный вручную, программа не меняет."""
        name, column = name.strip(), (date_column or "").strip()
        if not name:
            raise ValueError("Название не может быть пустым.")
        old = self.report(report_id)
        if column == (old["date_column"] or ""):
            source = old["date_source"]
        else:
            source = "manual" if column else "auto"
        try:
            self.db.execute("UPDATE reports SET name = ?, date_column = ?, date_source = ? "
                            "WHERE id = ?", (name, column or None, source, report_id))
            self.db.commit()
        except sqlite3.IntegrityError:
            raise ValueError(f"«{name}» уже есть в списке.") from None

    def set_found_column(self, report_id, column):
        self.db.execute("UPDATE reports SET date_column = ? WHERE id = ?", (column, report_id))

    def delete_system(self, system_id):
        try:
            self.db.execute("DELETE FROM systems WHERE id = ?", (system_id,))
            self.db.commit()
        except sqlite3.IntegrityError:
            self.db.rollback()
            raise ValueError("Нельзя удалить систему, пока у неё есть отчёты.") from None

    def delete_report(self, report_id):
        try:
            self.db.execute("DELETE FROM reports WHERE id = ?", (report_id,))
            self.db.commit()
        except sqlite3.IntegrityError:
            self.db.rollback()
            raise ValueError("Нельзя удалить отчёт, пока по нему есть сверки. "
                             "Сначала удалите их в истории.") from None

    # ---------- сверки ----------

    def create_run(self, report, file_a, file_b, period_from, period_to, options):
        """Запись о начатой сверке; до завершения она числится ошибкой."""
        cur = self.db.execute(
            "INSERT INTO runs (created_at, report_id, system_id, period_from, period_to, "
            "status, error_text, file_a, file_b, options) "
            "VALUES (?, ?, ?, ?, ?, 'error', 'Сверка не завершена', ?, ?, ?)",
            (datetime.now().isoformat(" ", "seconds"),
             report["id"] if report else None, report["system_id"] if report else None,
             period_from.isoformat() if period_from else "",
             period_to.isoformat() if period_to else "", str(file_a), str(file_b), options))
        self.db.commit()
        return cur.lastrowid

    def finish_run(self, run_id, result):
        columns = ", ".join(f"{k} = ?" for k in result)
        self.db.execute(f"UPDATE runs SET {columns} WHERE id = ?", (*result.values(), run_id))
        self.db.commit()

    def run(self, run_id):
        return self.db.execute(RUNS + " WHERE r.id = ?", (run_id,)).fetchone()

    def runs(self, system_id=None, report_id=None, status=None, limit=1000):
        """История, новые сверху. Возвращает (сверки, сколько всего найдено)."""
        where, params = [], []
        if report_id:
            where.append("r.report_id = ?")
            params.append(report_id)
        if system_id:
            where.append("r.system_id = ?")
            params.append(system_id)
        if status in ("ok", "diff", "error"):
            where.append("r.status = ?")
            params.append(status)
        clause = (" WHERE " + " AND ".join(where)) if where else ""
        total = self.db.execute("SELECT COUNT(*) FROM runs r" + clause, params).fetchone()[0]
        rows = self.db.execute(RUNS + clause + " ORDER BY r.id DESC LIMIT ?",
                               (*params, limit)).fetchall()
        return rows, total

    def delete_run(self, run_id):
        self.db.execute("DELETE FROM runs WHERE id = ?", (run_id,))
        self.db.commit()
        try:
            self.report_path(run_id).unlink()
        except OSError:
            pass

    def months(self, run_id):
        return self.db.execute("SELECT * FROM run_months WHERE run_id = ? ORDER BY month",
                               (run_id,)).fetchall()

    def diffs(self, run_id, month="", offset=0, limit=50):
        """
        Расхождения сверки: month - «ГГГГ-ММ», «none» (без месяца) или пусто (все).
        Возвращает (строки, сколько всего по этому условию).
        """
        where, params = "run_id = ?", [run_id]
        if month == "none":
            where += " AND month IS NULL"
        elif month:
            where += " AND month = ?"
            params.append(month)
        total = self.db.execute(f"SELECT COUNT(*) FROM run_diffs WHERE {where}",
                                params).fetchone()[0]
        rows = self.db.execute(f"SELECT * FROM run_diffs WHERE {where} ORDER BY id "
                               "LIMIT ? OFFSET ?", (*params, limit, offset)).fetchall()
        return rows, total
