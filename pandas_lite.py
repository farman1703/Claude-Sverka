"""
Замена библиотеки pandas для графического приложения - на чистом Python.

Зачем: pandas и numpy состоят из скомпилированных файлов (.pyd/.dll), которые
Windows (Smart App Control и похожие защиты) может не дать загрузить. Здесь
только Python-код и библиотека openpyxl (тоже чистый Python), блокировать нечего.

Что здесь есть: ровно те возможности pandas, которыми пользуется compare_excel.py
(чтение книги, таблица со столбцами, запись листов отчёта, разбор даты). Сверка
и построение отчёта остаются в compare_excel.py, этот модуль в них не участвует.
Поведение повторяет pandas 3 с движком openpyxl: те же значения ячеек при
чтении, те же типы и те же ячейки при записи.

Подключение (делает sverka_gui.py до импорта compare_excel):
    import pandas_lite
    sys.modules["pandas"] = pandas_lite
"""

import datetime as _dt
import re
import time as _time
import zipfile
from collections import defaultdict
from decimal import Decimal

__version__ = "lite"
nan = float("nan")
INT64_LIMIT = 2 ** 63


class _Missing:
    """Особые значения pandas «нет данных»; при чтении через openpyxl не встречаются."""

    def __init__(self, name):
        self._name = name

    def __repr__(self):
        return self._name


NaT = _Missing("NaT")
NA = _Missing("<NA>")


def _isna(v):
    return (v is None or v is NaT or v is NA or (isinstance(v, float) and v != v)
            or (isinstance(v, Decimal) and v.is_nan()))


def _infer(values):
    """
    Как pandas подбирает тип столбца: если в столбце только числа, то целые рядом
    с дробными или с пропусками становятся дробными (5 -> 5.0), пропуски - NaN.
    Столбцы с текстом, датами, логическими значениями остаются как есть.
    """
    ints = floats = nulls = False
    for v in values:
        if v is None:
            nulls = True
        elif isinstance(v, bool):
            return values
        elif isinstance(v, int):
            if abs(v) >= INT64_LIMIT:
                return values
            ints = True
        elif isinstance(v, float):
            floats = True
        else:
            return values
    if not (ints or floats) or not (floats or nulls):
        return values
    return [nan if v is None else float(v) for v in values]


def _labels(names):
    """
    Названия столбцов, как их хранит pandas: если все они - числа и есть дробное,
    целые становятся дробными; среди одних дат ячейка с ошибкой становится NaT.
    """
    names = list(names)
    if not names:
        return names
    empty = [isinstance(c, float) and c != c for c in names]
    if all(isinstance(c, (int, float)) and not isinstance(c, bool) for c in names):
        if any(isinstance(c, float) for c in names):
            return [float(c) for c in names]
    elif any(empty) and all(e or isinstance(c, _dt.datetime) for c, e in zip(names, empty)):
        return [NaT if e else c for c, e in zip(names, empty)]
    return names


# ---------- таблица ----------

class Series:
    """Один столбец таблицы."""

    def __init__(self, values):
        self._values = list(values)

    def __len__(self):
        return len(self._values)

    def __iter__(self):
        return iter(self._values)

    def tolist(self):
        return list(self._values)

    def map(self, arg):
        if isinstance(arg, dict):
            return Series(_infer([arg.get(v, nan) for v in self._values]))
        return Series(_infer([arg(v) for v in self._values]))

    def __eq__(self, other):
        return Series([v == other for v in self._values])

    __hash__ = None

    def all(self):
        return all(self._values)


class _RowSetter:
    def __init__(self, frame):
        self._frame = frame

    def __setitem__(self, index, values):
        frame, values = self._frame, list(values)
        if len(values) != len(frame._data):
            raise ValueError("cannot set a row with mismatched columns")
        if index == frame._nrows:
            for column, value in zip(frame._data, values):
                column.append(value)
            frame._nrows += 1
        elif isinstance(index, int) and 0 <= index < frame._nrows:
            for column, value in zip(frame._data, values):
                column[index] = value
        else:
            raise KeyError(index)


class DataFrame:
    """Таблица: список названий столбцов и значения по столбцам."""

    def __init__(self, data=None, columns=None):
        rows = list(data) if data is not None else []
        if rows and all(isinstance(r, dict) for r in rows):
            names = (list(columns) if columns is not None
                     else list(dict.fromkeys(k for r in rows for k in r)))
            cols = [_infer([r.get(name, nan) for r in rows]) for name in names]
        elif rows:
            width = max(len(r) for r in rows)
            names = list(columns) if columns is not None else list(range(width))
            if len(names) != width:
                raise ValueError(f"{len(names)} columns passed, passed data had {width} columns")
            cols = [_infer([r[k] if k < len(r) else None for r in rows]) for k in range(width)]
        else:
            names = list(columns) if columns is not None else []
            cols = [[] for _ in names]
        self._columns, self._data, self._nrows = names, cols, len(rows)

    @classmethod
    def _make(cls, names, cols, nrows):
        """Готовая таблица без подбора типов (так читаются листы: dtype=object)."""
        frame = cls()
        frame._columns, frame._data, frame._nrows = list(names), cols, nrows
        return frame

    @property
    def columns(self):
        return self._columns

    @property
    def empty(self):
        return not self._nrows or not self._columns

    @property
    def loc(self):
        return _RowSetter(self)

    def __len__(self):
        return self._nrows

    def __contains__(self, name):
        return name in self._columns

    def _position(self, name):
        try:
            return self._columns.index(name)
        except ValueError:
            raise KeyError(name) from None

    def __getitem__(self, name):
        return Series(self._data[self._position(name)])

    def __setitem__(self, name, values):
        values = values.tolist() if isinstance(values, Series) else list(values)
        if name in self._columns:
            self._data[self._position(name)] = values
        else:
            self._columns.append(name)
            self._data.append(values)
            self._nrows = len(values)

    def rename(self, columns):
        return DataFrame._make([columns.get(c, c) for c in self._columns],
                               [list(col) for col in self._data], self._nrows)

    def drop(self, columns):
        names = list(columns) if isinstance(columns, (list, tuple, set)) else [columns]
        for name in names:
            self._position(name)   # нет такого столбца - ошибка, как в pandas
        keep = [k for k, c in enumerate(self._columns) if c not in names]
        return DataFrame._make([self._columns[k] for k in keep],
                               [list(self._data[k]) for k in keep], self._nrows)

    def set_axis(self, labels, axis=0):
        labels = list(labels)
        if axis != 1:
            raise NotImplementedError("set_axis: поддерживается только axis=1")
        if len(labels) != len(self._columns):
            raise ValueError(f"Length mismatch: Expected axis has {len(self._columns)} "
                             f"elements, new values have {len(labels)} elements")
        return DataFrame._make(_labels(labels), [list(col) for col in self._data], self._nrows)

    def apply(self, func):
        return DataFrame._make(self._columns,
                               [func(Series(col)).tolist() for col in self._data], self._nrows)

    def to_excel(self, writer, sheet_name="Sheet1", index=True, header=True, startrow=0):
        """Пишет таблицу на лист так же, как pandas: сначала заголовки, затем столбцы."""
        if index:
            raise NotImplementedError("to_excel: поддерживается только index=False")
        ws = writer._sheet(sheet_name)
        row = startrow + 1
        if header:
            for k, name in enumerate(self._columns, 1):
                _write_cell(ws, row, k, name)
            row += 1
        for k, column in enumerate(self._data, 1):
            for i, value in enumerate(column):
                _write_cell(ws, row + i, k, value)


def _write_cell(ws, row, col, val):
    """Одна ячейка: пропуск пишется пустой строкой, значения - как их пишет pandas."""
    fmt = None
    if _isna(val):
        val = ""
    elif isinstance(val, float) and val in (float("inf"), float("-inf")):
        val = "inf" if val > 0 else "-inf"
    if getattr(val, "tzinfo", None) is not None:
        raise ValueError("Excel does not support datetimes with timezones. Please ensure "
                         "that datetimes are timezone unaware before writing to Excel.")
    if isinstance(val, (bool, int, float, Decimal)):
        pass
    elif isinstance(val, _dt.datetime):
        fmt = "YYYY-MM-DD HH:MM:SS"
    elif isinstance(val, _dt.date):
        fmt = "YYYY-MM-DD"
    elif isinstance(val, _dt.timedelta):
        val, fmt = val.total_seconds() / 86400, "0"
    else:
        val = str(val)
    cell = ws.cell(row=row, column=col)
    cell.value = val
    if fmt:
        cell.number_format = fmt


class ExcelWriter:
    """Книга Excel, в которую таблицы пишут свои листы; сохраняется при выходе из with."""

    def __init__(self, path, engine=None, **kwargs):
        from openpyxl import Workbook
        if engine not in (None, "openpyxl"):
            raise ValueError(f"No Excel writer '{engine}'")
        self._handle = open(path, "wb")   # как в pandas: файл занят - ошибка сразу
        self.book = Workbook()
        if self.book.worksheets:
            self.book.remove(self.book.worksheets[0])

    @property
    def sheets(self):
        return {ws.title: ws for ws in self.book.worksheets}

    def _sheet(self, name):
        sheets = self.sheets
        if name in sheets:
            return sheets[name]
        ws = self.book.create_sheet()
        ws.title = name
        return ws

    def close(self):
        try:
            self.book.save(self._handle)
        finally:
            self._handle.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()


# ---------- чтение книги ----------

XLS_SIGNATURES = (b"\x09\x00\x04\x00\x07\x00\x10\x00", b"\x09\x02\x06\x00\x00\x00\x10\x00",
                  b"\x09\x04\x06\x00\x00\x00\x10\x00", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")


NOT_EXCEL = ("Это не файл Excel или он повреждён. Нужен файл в формате .xlsx: откройте его "
             "в Excel и сохраните через «Файл → Сохранить как → Книга Excel».")


def _check_format(handle):
    """Определяет формат файла по содержимому, как pandas; читаем только xlsx/xlsm."""
    handle.seek(0)
    peek = handle.read(8)
    handle.seek(0)
    if any(peek.startswith(sig) for sig in XLS_SIGNATURES):
        raise ValueError("Файл в старом формате .xls. Откройте его в Excel и сохраните "
                         "как .xlsx")
    if not peek.startswith(b"PK\x03\x04"):
        raise ValueError(NOT_EXCEL)
    with zipfile.ZipFile(handle) as archive:
        names = [name.replace("\\", "/").lower() for name in archive.namelist()]
    handle.seek(0)
    if "xl/workbook.xml" in names:
        return
    if "xl/workbook.bin" in names:
        raise ValueError("Файл в формате .xlsb. Откройте его в Excel и сохраните как .xlsx")
    if "content.xml" in names:
        raise ValueError("Файл в формате .ods. Откройте его в Excel и сохраните как .xlsx")
    raise ValueError(NOT_EXCEL)


def _sheet_rows(sheet):
    """Значения листа построчно - в точности как их отдаёт pandas с движком openpyxl."""
    from openpyxl.cell.cell import TYPE_ERROR, TYPE_NUMERIC

    def convert(cell):
        value = cell.value
        if value is None:
            return ""
        if cell.data_type == TYPE_ERROR:   # #Н/Д, #ДЕЛ/0! и т.п.
            return nan
        if cell.data_type == TYPE_NUMERIC:
            whole = int(value)
            return whole if whole == value else float(value)
        return value

    sheet.reset_dimensions()   # размеры, записанные в файле, бывают неверными
    data, last = [], -1
    for number, row in enumerate(sheet.rows):
        values = [convert(cell) for cell in row]
        while values and values[-1] == "":
            values.pop()
        if values:
            last = number
        data.append(values)
    data = data[:last + 1]
    if data:
        width = max(len(r) for r in data)
        data = [r + [""] * (width - len(r)) for r in data]
    return data


def _header_names(line):
    """Названия столбцов из строки заголовков: пустые - «Unnamed: N», повторы - «имя.1»."""
    names, unnamed = [], []
    for i, c in enumerate(line):
        if c == "":
            unnamed.append(i)
            names.append(f"Unnamed: {i}")
        else:
            names.append(c)
    counts = defaultdict(int)
    for i in [i for i in range(len(names)) if i not in unnamed] + unnamed:
        col = old = names[i]
        count = counts[col]
        if count > 0:
            while count > 0:
                counts[old] = count + 1
                col = f"{old}.{count}"
                if col in names:
                    count += 1
                else:
                    count = counts[col]
        names[i] = col
        counts[col] = count + 1
    return _labels(names)


def read_excel(io, sheet_name=0, header=0, dtype=None, engine=None,
               keep_default_na=True, na_filter=True):
    """Читает все листы книги: {название листа: таблица}. Значения ячеек не меняются."""
    if engine not in (None, "openpyxl"):
        raise ValueError(f"Unknown engine: {engine}")
    if (sheet_name is not None or dtype is not object or na_filter or keep_default_na
            or not isinstance(header, int) or header < 0):
        raise NotImplementedError("read_excel: поддерживается только вызов из compare_excel.py")
    from openpyxl import load_workbook

    result = {}
    with open(io, "rb") as handle:
        _check_format(handle)
        book = load_workbook(handle, read_only=True, data_only=True, keep_links=False)
        try:
            for sheet in book.worksheets:
                data = _sheet_rows(sheet)
                if not data:
                    result[sheet.title] = DataFrame()
                    continue
                if header >= len(data):
                    raise ValueError(f"Passed header=[{header}], len of 1, but only "
                                     f"{len(data)} lines in file (sheet: {sheet.title})")
                names = _header_names(data[header])
                rows = data[header + 1:]
                cols = [[r[k] for r in rows] for k in range(len(names))]
                result[sheet.title] = DataFrame._make(names, cols, len(rows))
        finally:
            book.close()
    return result


# ---------- разбор даты ----------
# Нужен только для пояснения к техническим столбцам («похоже на время формирования
# отчёта, разница между файлами ...»). Правила - как у pandas.to_datetime.

DELIMITED_RE = re.compile(r"(\d{1,2})[/\-. ](\d{1,2})[/\-. ](\d{4})")
MONTH_YEAR_RE = re.compile(r"(\d{1,2})([/\- ])(\d{4})")
QUARTER_RE = re.compile(r"(?:(\d{4}|\d\d)-?[Qq]([1-4])|([1-4])[Qq]-?(\d{4}|\d\d))")
TIME_RE = re.compile(r"(\d{1,2}):(\d\d)")
NOT_DATELIKE = {"a", "A", "m", "M", "p", "P", "t", "T"}
NAT_TEXTS = {"NaT", "nat", "NAT", "nan", "NaN", "NAN"}
START = _dt.datetime(1, 1, 1)


class _NoDate:
    """«Не дата» (текст nan, NaT): разница с любой датой - тоже «не число», как в pandas."""

    def __sub__(self, other):
        return self

    __rsub__ = __sub__

    def total_seconds(self):
        return nan

    def __repr__(self):
        return "NaT"


class _BeforeEra:
    """Дата с годом меньше 1 (pandas читает текст «-5» как 5-й год до н.э.)."""

    def __init__(self, year, month, day):
        # число дней от 0001-01-01 по григорианскому календарю, продолженному в прошлое
        y = year - (month <= 2)
        era = y // 400
        yoe = y - era * 400
        doy = (153 * (month + (-3 if month > 2 else 9)) + 2) // 5 + day - 1
        self.days = era * 146097 + yoe * 365 + yoe // 4 - yoe // 100 + doy - 306

    def _delta(self, other):
        if isinstance(other, _BeforeEra):
            return _dt.timedelta(days=self.days - other.days)
        if isinstance(other, _dt.datetime) and other.tzinfo is None:
            return _dt.timedelta(days=self.days) - (other - START)
        return NotImplemented

    def __sub__(self, other):
        return self._delta(other)

    def __rsub__(self, other):
        delta = self._delta(other)
        return delta if delta is NotImplemented else -delta


def _iso(text):
    """
    Дата в записи ISO 8601 в том виде, как её принимает pandas: 2026-11-13,
    2026.5, 2026/11/13 7:05, 20261113T100000, 2026-11-13T10:00:00.5+03:00.
    Год меньше 1 («-5», «.5») возвращается особым объектом. Если текст так
    не записан - None.
    """
    s = text.lstrip()
    n = len(s)
    negative = text.startswith("-")

    def digit(k):
        return k < n and "0" <= s[k] <= "9"

    def number(k, optional_second):
        """Одна-две цифры с позиции k: (значение, следующая позиция) или None."""
        if not digit(k):
            return None
        if digit(k + 1):
            return int(s[k:k + 2]), k + 2
        return (int(s[k]), k + 1) if optional_second else None

    i = j = 1 if negative else 0
    while digit(j) and j - i < 4:
        j += 1
    if j == i:                                  # года нет: «.5» - это май нулевого года
        if negative or i >= n or s[i] not in "./\\":
            return None
    elif not negative and j - i != 4:
        return None
    year, i = int(s[i:j] or 0), j
    month = day = 1
    hour = minute = second = micro = 0
    zone = None
    has_time = False
    if i < n:
        sep = None
        if not digit(i):                        # разделитель года и месяца
            if s[i] not in "-./\\ ":
                return None
            sep, i = s[i], i + 1
        found = number(i, sep is not None)
        if not found or not 1 <= found[0] <= 12:
            return None
        month, i = found
        if i == n:
            if sep is None:                     # ГГГГММ без дня датой не считается
                return None
        else:
            if sep is not None:
                if s[i] != sep:
                    return None
                i += 1
            found = number(i, sep is not None)
            if not found:
                return None
            day, i = found
            if i < n:                           # время
                if s[i] not in "T " or i == n - 1:
                    return None
                has_time = True
                found = number(i + 1, True)
                if not found or found[0] >= 24:
                    return None
                two_digits = digit(i + 2)
                hour, i = found
                colon = False
                more = i < n
                if more and s[i] == ":":
                    colon, i = True, i + 1
                    if not digit(i):
                        return None
                elif more and not digit(i):
                    if not two_digits:
                        return None
                    more = False
                elif not more and not two_digits:
                    return None
                if more:
                    found = number(i, colon)
                    if not found or found[0] >= 60:
                        return None
                    minute, i = found
                    seconds = False
                    if i < n and colon and s[i] == ":":
                        i += 1
                        if not digit(i):
                            return None
                        seconds = True
                    elif i < n and not colon and digit(i):
                        seconds = True
                    if seconds:
                        found = number(i, colon)
                        if not found or found[0] >= 60:
                            return None
                        second, i = found
                        if i < n and s[i] == ".":
                            j = i + 1
                            while digit(j):
                                j += 1
                            micro, i = int(s[i + 1:j][:6].ljust(6, "0") or 0), j
                while i < n and s[i].isspace():
                    i += 1
                if i < n and s[i] == "Z":
                    zone, i = _dt.timezone.utc, i + 1
                elif i < n and s[i] in "+-":
                    if not (digit(i + 1) and digit(i + 2)) or int(s[i + 1:i + 3]) >= 24:
                        return None
                    sign, hours, minutes, i = s[i], int(s[i + 1:i + 3]), 0, i + 3
                    if i < n:
                        if s[i] == ":":
                            i += 1
                        if not (digit(i) and digit(i + 1)) or int(s[i:i + 2]) >= 60:
                            return None
                        minutes, i = int(s[i:i + 2]), i + 2
                    offset = _dt.timedelta(hours=hours, minutes=minutes)
                    zone = _dt.timezone(-offset if sign == "-" else offset)
                while i < n and s[i].isspace():
                    i += 1
                if i != n:
                    return None
    try:
        if negative or year == 0:
            _dt.datetime(2000 - year % 400, month, day)      # проверка числа месяца
            if has_time or (negative and s != text):
                return None
            return _BeforeEra(-year, month, day)
        return _dt.datetime(year, month, day, hour, minute, second, micro, zone)
    except ValueError:
        return None


def _free_form(text, default, need_day=False):
    """Дата в свободной записи («13 Nov 2026 10:00») через dateutil - так же, как в pandas."""
    try:
        from dateutil.parser import DEFAULTPARSER
    except ImportError:
        raise ValueError(f"Unknown datetime string format: {text}") from None
    try:
        found, _ = DEFAULTPARSER._parse(text, dayfirst=False, yearfirst=False)
    except OverflowError as e:
        raise ValueError(f"Unknown datetime string format: {text}") from e
    if found is None:
        raise ValueError(f"Unknown datetime string format, unable to parse: {text}")
    parts = {name: getattr(found, name) for name in
             ("year", "month", "day", "hour", "minute", "second", "microsecond")
             if getattr(found, name) is not None}
    if not parts or (need_day and not {"year", "month", "day"} <= set(parts)):
        raise ValueError(f"Unable to parse datetime string: {text}")
    result = default.replace(**parts)
    if found.weekday is not None and not found.day:
        raise ValueError("Parsing datetimes with weekday but no day information is not supported")
    if found.tzname and found.tzname in _time.tzname:
        if found.tzname != "UTC":
            raise ValueError(f"Parsing '{found.tzname}' as tzlocal is not supported")
        result = result.replace(tzinfo=_dt.timezone.utc)
    elif found.tzoffset == 0:
        result = result.replace(tzinfo=_dt.timezone.utc)
    elif found.tzoffset:
        if not -86400 < found.tzoffset < 86400:
            raise ValueError(f"offset must be a timedelta strictly between -24h and 24h: {text}")
        result = result.replace(tzinfo=_dt.timezone(_dt.timedelta(seconds=found.tzoffset)))
    elif found.tzname is not None:
        raise ValueError(f'Parsed string "{text}" included an un-recognized timezone '
                         f'"{found.tzname}"')
    return result


def _parse_text(text):
    """Дата из текста по правилам pandas (месяц раньше дня, если это возможно)."""
    if text in NAT_TEXTS or text == "":
        return _NoDate()
    if text.startswith("-"):
        # Полная дата с четырёхзначным годом читается без минуса; иначе число
        # после минуса - это год до нашей эры («-5», «-5.11»); остальное - не дата
        if re.search(r"\d{4}", text):
            try:
                found = _free_form(text, START, need_day=True)
            except (ValueError, OverflowError):
                found = None
            if found is not None:
                return found
        found = _iso(text)
        if found is None:
            raise ValueError(f"Unknown datetime string format, unable to parse: {text}")
        return found
    found = _iso(text)
    if found is not None:
        return found
    # Текст, похожий на обычное число (10, 1.5), датой не считается
    if text[:1] != "0":
        if text in NOT_DATELIKE:
            raise ValueError(f'Given date string "{text}" not likely a datetime')
        try:
            number = float(text)
        except ValueError:
            pass
        else:
            if not number >= 1000:
                raise ValueError(f'Given date string "{text}" not likely a datetime')
    m = TIME_RE.match(text)
    if m and len(text) >= 4 and int(m[1]) <= 23 and int(m[2]) <= 59:
        # только время: дата - сегодняшняя
        return _free_form(text, _dt.datetime.combine(_dt.date.today(), _dt.time()))
    m = DELIMITED_RE.fullmatch(text)
    if m:
        month, day, year = int(m[1]), int(m[2]), int(m[3])
        if month > 12:
            month, day = day, month
        return _dt.datetime(year, month, day)
    m = MONTH_YEAR_RE.fullmatch(text)
    if m:
        return _dt.datetime(int(m[3]), int(m[1]), 1)
    m = QUARTER_RE.fullmatch(text)
    if m:
        year, quarter = (m[1], m[2]) if m[1] else (m[4], m[3])
        year = int(year) + (2000 if len(year) == 2 else 0)
        return _dt.datetime(year, 3 * int(quarter) - 2, 1)
    return _free_form(text, START)


def to_datetime(value, errors="raise"):
    """Дата и время из значения ячейки; если это не дата - ошибка."""
    if isinstance(value, _dt.datetime):
        return value
    if isinstance(value, _dt.date):
        return _dt.datetime(value.year, value.month, value.day)
    if isinstance(value, str):
        return _parse_text(value)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return START       # число: вызывающий код такое значение датой не считает
    raise TypeError(f"{type(value)} is not convertible to datetime")
