#!/usr/bin/env python3
"""
Собирает архив для установки приложения на компьютер: dist/SverkaExcel-setup.zip

Запуск:
    python make_zip.py
"""

import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FOLDER = "SverkaExcel"                      # папка внутри архива
TARGET = ROOT / "dist" / "SverkaExcel-setup.zip"
FILES = ["compare_excel.py", "pandas_lite.py", "engine.py", "storage.py", "service.py",
         "sverka_gui.py", "gui_widgets.py", "gui_result.py", "gui_history.py", "gui_dicts.py",
         "install.py", "requirements.txt",
         "install.bat", "run.bat", "install.command", "install.sh", "run.sh"]
EXECUTABLE = (".py", ".sh", ".command")


def add(archive, name, data, executable=False):
    info = zipfile.ZipInfo(f"{FOLDER}/{name}", date_time=(2026, 1, 1, 0, 0, 0))
    info.compress_type = zipfile.ZIP_DEFLATED
    info.create_system = 3                  # права доступа в формате Unix
    info.external_attr = (0o755 if executable else 0o644) << 16
    archive.writestr(info, data)


def main():
    TARGET.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(TARGET, "w") as archive:
        for name in FILES:
            data = (ROOT / name).read_bytes()
            if name.endswith(".bat"):       # Windows ждёт в .bat переводы строк CRLF
                data = data.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
            elif name.endswith((".sh", ".command")):
                data = data.replace(b"\r\n", b"\n")
            add(archive, name, data, executable=name.endswith(EXECUTABLE))
        # Инструкция обычным текстом; BOM - чтобы Блокнот Windows верно показал русский
        readme = (ROOT / "README.md").read_text(encoding="utf-8")
        readme = "\n".join(line for line in readme.split("\n")      # без картинок
                           if not line.startswith("![")).replace("\n\n\n", "\n\n")
        add(archive, "README.txt", b"\xef\xbb\xbf" + readme.replace("\n", "\r\n").encode("utf-8"))
    print(f"Готово: {TARGET} ({TARGET.stat().st_size // 1024} КБ)")


if __name__ == "__main__":
    main()
