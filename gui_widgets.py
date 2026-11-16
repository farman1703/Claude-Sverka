"""Общие элементы окна: цвета, шрифты, прокручиваемая область, цветная таблица."""

import os
import subprocess
import sys
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

BG, CARD, LINE = "#F5F5F2", "#FFFFFF", "#D5D4CE"
TEXT, MUTED = "#1D1D1B", "#6B6A66"
OK_BG, OK_FG = "#C6EFCE", "#006100"
BAD_BG, BAD_FG = "#FFC7CE", "#9C0006"
WARN_BG, WARN_FG = "#FFEB9C", "#7F6000"
GREY_BG, HEAD_BG = "#EDEDED", "#305496"
PILLS = {"ok": (OK_BG, OK_FG), "bad": (BAD_BG, BAD_FG), "warn": (WARN_BG, WARN_FG),
         "err": (GREY_BG, MUTED), "plain": (GREY_BG, TEXT)}

# Ячейки таблицы расхождений: фон, цвет текста, шрифт. Цвета - как в Excel-отчёте
CELLS = {
    "head": (HEAD_BG, "#FFFFFF", "bold"), "side": (GREY_BG, TEXT, "bold"),
    "plain": (CARD, TEXT, "base"), "same": (CARD, "#8A8A85", "base"),
    "x": ("#FFD966", TEXT, "bold"),          # значение есть в обоих файлах, но разное
    "have": (WARN_BG, TEXT, "bold"),         # значение есть только в этом файле
    "e": (BAD_BG, BAD_FG, "italic"),         # (пусто): значения нет, а в другом файле есть
    "tech": (GREY_BG, "#7F7F7F", "base"),    # технический столбец
    "only_a": ("#F8CBAD", TEXT, "base"), "only_b": ("#BDD7EE", TEXT, "base"),
}


def make_fonts(root):
    """Шрифты окна на основе системного."""
    base = tkfont.nametofont("TkDefaultFont")
    size = abs(base.cget("size")) or 10
    fonts = {"base": base}
    for name, options in (("bold", {"weight": "bold"}), ("italic", {"slant": "italic"}),
                          ("small", {"size": max(size - 1, 8)}),
                          ("head", {"size": size + 2, "weight": "bold"}),
                          ("title", {"size": size + 5, "weight": "bold"}),
                          ("big", {"size": size + 6, "weight": "bold"})):
        fonts[name] = base.copy()
        fonts[name].configure(**options)
    return fonts


def open_path(path):
    """Открывает файл или папку программой по умолчанию."""
    if sys.platform == "win32":
        os.startfile(path)   # noqa: S606 - открываем то, что сами создали
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def fix_tree_colors(style):
    """В части версий Tk строки таблицы не красятся по меткам; это известное исправление."""
    def kept(option):
        return [e for e in style.map("Treeview", query_opt=option)
                if e[:2] != ("!disabled", "!selected")]
    style.map("Treeview", foreground=kept("foreground"), background=kept("background"))


def pill(parent, text, kind, fonts, **options):
    """Цветная метка: «совпало», «3 расхождения», «нет в файле 2»."""
    bg, fg = PILLS[kind]
    options = {"padx": 8, "pady": 2, **options}
    return tk.Label(parent, text=text, bg=bg, fg=fg, font=fonts["small"], **options)


def clear(frame):
    for child in frame.winfo_children():
        child.destroy()


def install_wheel(root):
    """
    Прокрутка колесом мыши области под курсором. Области отмечают себя сами:
    wheel_y - что прокручивать по вертикали, wheel_x - по горизонтали (с Shift).
    """
    native = ("Text", "Treeview", "Listbox", "TCombobox", "TSpinbox")

    def target(event, attribute):
        try:
            widget = root.winfo_containing(event.x_root, event.y_root)
        except (KeyError, tk.TclError):
            return None
        while widget is not None:
            if attribute == "wheel_y" and widget.winfo_class() in native:
                return None                 # такие элементы прокручиваются сами
            if getattr(widget, attribute, None) is not None:
                return getattr(widget, attribute)
            widget = widget.master
        return None

    def scroll(event, attribute):
        canvas = target(event, attribute)
        if canvas is None:
            return
        view = canvas.yview() if attribute == "wheel_y" else canvas.xview()
        if view == (0.0, 1.0):
            return
        up = event.num == 4 or event.delta > 0
        steps = max(1, abs(event.delta) // 120) * 3
        move = canvas.yview_scroll if attribute == "wheel_y" else canvas.xview_scroll
        move(-steps if up else steps, "units")

    for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
        root.bind_all(sequence, lambda e: scroll(e, "wheel_y"), add="+")
    for sequence in ("<Shift-MouseWheel>", "<Shift-Button-4>", "<Shift-Button-5>"):
        root.bind_all(sequence, lambda e: scroll(e, "wheel_x"))


class Scrolled(tk.Frame):
    """Область с вертикальной прокруткой. Содержимое кладётся в .body."""

    def __init__(self, parent, bg=BG):
        super().__init__(parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, yscrollincrement=20)
        bar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.body = tk.Frame(self.canvas, bg=bg)
        window = self.canvas.create_window(0, 0, anchor="nw", window=self.body)
        self.body.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=(0, 0, e.width, e.height)))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(window, width=e.width))
        self.canvas.wheel_y = self.canvas

    def to_top(self):
        self.canvas.yview_moveto(0)


class GridTable(tk.Frame):
    """
    Таблица с цветными ячейками, как лист Excel. Полоса горизонтальной прокрутки
    стоит сверху, чтобы она была под рукой у длинной таблицы; с Shift работает колесо.
    """

    ROW, PAD, MAX_WIDTH = 24, 8, 260

    def __init__(self, parent, fonts):
        super().__init__(parent, bg=BG)
        self.fonts = fonts
        self.canvas = tk.Canvas(self, bg=CARD, highlightthickness=1, highlightbackground=LINE,
                                xscrollincrement=20, height=10)
        self.bar = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=self.bar.set)
        self.bar.pack(fill="x")
        self.canvas.pack(fill="x")
        self.canvas.wheel_x = self.canvas

    def _fit(self, text, font, width):
        """Текст, обрезанный под ширину ячейки."""
        room = width - 2 * self.PAD
        measured = font.measure(text)
        if measured <= room:
            return text
        return text[:max(1, int(len(text) * room / measured) - 1)] + "…"

    def show(self, headers, rows, last_wide=True):
        """
        headers - названия столбцов (можно в две строки через \\n);
        rows - строки: список ячеек (текст, вид из CELLS, сколько строк занимает).
        Ячейка None пропускается (её место занято ячейкой выше).
        """
        canvas, fonts = self.canvas, self.fonts
        canvas.delete("all")
        ncols = len(headers)
        widths = []
        for c in range(ncols):
            wide = last_wide and c == ncols - 1
            limit = 420 if wide else self.MAX_WIDTH
            need = max(fonts["bold"].measure(line) for line in headers[c].split("\n"))
            for row in rows:
                cell = row[c] if c < len(row) else None
                if cell and cell[0]:
                    need = max(need, fonts[CELLS[cell[1]][2]].measure(cell[0]))
                    if need >= limit:
                        break
            widths.append(max(44, min(need + 2 * self.PAD + 2, limit)))
        head_height = 40 if any("\n" in h for h in headers) else 26
        x = 0
        for c, title in enumerate(headers):
            bg, fg, font = CELLS["head"]
            canvas.create_rectangle(x, 0, x + widths[c], head_height, fill=bg, outline=LINE)
            canvas.create_text(x + self.PAD, head_height / 2, anchor="w", fill=fg,
                               font=fonts[font], text="\n".join(
                                   self._fit(line, fonts[font], widths[c])
                                   for line in title.split("\n")))
            x += widths[c]
        y = head_height
        for row in rows:
            x = 0
            for c in range(ncols):
                cell = row[c] if c < len(row) else ("", "plain", 1)
                if cell is not None:
                    text, kind, span = cell
                    bg, fg, font = CELLS[kind]
                    height = self.ROW * span
                    canvas.create_rectangle(x, y, x + widths[c], y + height, fill=bg, outline=LINE)
                    if text:
                        if span > 1:        # объединённая ячейка: текст с переносом
                            canvas.create_text(x + self.PAD, y + height / 2, anchor="w", fill=fg,
                                               font=fonts[font], text=text,
                                               width=widths[c] - 2 * self.PAD)
                        else:
                            canvas.create_text(x + self.PAD, y + height / 2, anchor="w", fill=fg,
                                               font=fonts[font],
                                               text=self._fit(text, fonts[font], widths[c]))
                x += widths[c]
            y += self.ROW
        canvas.configure(height=y + 1, scrollregion=(0, 0, sum(widths) + 1, y + 1))
        canvas.xview_moveto(0)
