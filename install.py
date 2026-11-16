#!/usr/bin/env python3
"""
Установка приложения «Сверка Excel» на компьютер (Windows, macOS, Linux).

Что делает:
    1. создаёт в папке приложения отдельное окружение Python (.venv);
    2. ставит в него библиотеки из requirements.txt (нужен интернет);
    3. создаёт ярлык для запуска на рабочем столе.

Ставятся только библиотеки на чистом Python (openpyxl, python-dateutil): в них
нет скомпилированных файлов, которые Windows может заблокировать при загрузке.

Запускается файлами install.bat (Windows), install.command (macOS) или
install.sh (Linux). Повторный запуск безопасен: он обновляет установку.
Папку приложения после установки не перемещайте - иначе запустите установку снова.
"""

import os
import plistlib
import shutil
import stat
import subprocess
import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
VENV = APP_DIR / ".venv"
GUI = APP_DIR / "sverka_gui.py"
APP_NAME = "Сверка Excel"
REQUIRED = ["openpyxl"]                    # python-dateutil нужна только для одного пояснения
BLOCKED = ("pandas", "numpy", "python_calamine")   # скомпилированные, приложению не нужны
WINDOWS = sys.platform == "win32"
MACOS = sys.platform == "darwin"


def fail(message):
    print("\nОШИБКА: " + message)
    sys.exit(1)


def venv_python(windowed=False):
    if WINDOWS:
        return VENV / "Scripts" / ("pythonw.exe" if windowed else "python.exe")
    return VENV / "bin" / "python"


def works(python, code):
    if not python.exists():
        return False
    return subprocess.run([str(python), "-c", code], capture_output=True).returncode == 0


def check_python():
    if sys.version_info < (3, 9):
        fail(f"нужен Python 3.9 или новее, а запущен {sys.version.split()[0]}. "
             "Установите свежий Python с https://www.python.org/downloads/")
    try:
        import tkinter  # noqa: F401
    except ImportError:
        if MACOS:
            how = ("установите Python с https://www.python.org/downloads/ "
                   "(или выполните: brew install python-tk)")
        elif WINDOWS:
            how = ("переустановите Python с https://www.python.org/downloads/, "
                   "оставив включённым пункт «tcl/tk and IDLE»")
        else:
            how = ("выполните: sudo apt install python3-tk   "
                   "(Fedora: sudo dnf install python3-tkinter)")
        fail("в этом Python нет модуля tkinter, без него окно не откроется. Чтобы "
             "исправить, " + how)


def create_venv():
    if works(venv_python(), "import sys"):
        print("Окружение уже есть, обновляю библиотеки...")
        return
    print("Создаю окружение Python...")
    shutil.rmtree(VENV, ignore_errors=True)
    done = subprocess.run([sys.executable, "-m", "venv", str(VENV)])
    if done.returncode or not venv_python().exists():
        hint = "" if WINDOWS or MACOS else " Выполните: sudo apt install python3-venv"
        fail("не удалось создать окружение Python." + hint)


def install_packages():
    print("Устанавливаю библиотеки (нужен интернет)...")
    pip = [str(venv_python()), "-m", "pip", "install", "--disable-pip-version-check"]
    remove_compiled_packages()
    if subprocess.run(pip + ["-r", str(APP_DIR / "requirements.txt")]).returncode:
        print("\nНе всё установилось, ставлю только обязательные библиотеки...")
        if subprocess.run(pip + REQUIRED).returncode:
            fail("не удалось установить библиотеки. Проверьте подключение к интернету "
                 "и запустите установку снова.")
    check_imports()


def remove_compiled_packages():
    """
    Убирает pandas, numpy и python-calamine, оставшиеся от прежней версии приложения:
    их скомпилированные файлы Windows может блокировать, а приложению они не нужны.
    """
    code = ("import importlib.util as u, sys; "
            f"sys.exit(not any(u.find_spec(m) for m in {BLOCKED!r}))")
    if works(venv_python(), code):
        print("Удаляю библиотеки прежней версии (pandas, numpy)...")
        subprocess.run([str(venv_python()), "-m", "pip", "uninstall", "-y", "-q",
                        "pandas", "numpy", "python-calamine"])


def check_imports():
    """Проверяет, что приложение и библиотеки загружаются; если нет - показывает ошибку."""
    checks = (("openpyxl", "import openpyxl"), ("tkinter", "import tkinter"),
              ("программа", "import sys, sverka_gui; "
                            "sys.exit(str(sverka_gui.IMPORT_ERROR) if sverka_gui.IMPORT_ERROR else 0)"))
    failed = []
    for module, code in checks:
        done = subprocess.run([str(venv_python()), "-c", code], cwd=str(APP_DIR),
                              capture_output=True, text=True, errors="replace")
        if done.returncode:
            text = (done.stderr or done.stdout or "").strip()
            code = done.returncode
            failed.append((module, text or "сообщения нет, код завершения "
                                           f"{code} (0x{code & 0xFFFFFFFF:08X})"))
    if not failed:
        return
    for module, text in failed:
        print(f"\n--- не загружается {module} ---\n{text}")
    everything = " ".join(text for _, text in failed)
    hints = []
    if "Application Control" in everything or "0xC0E90002" in everything.upper():
        hints.append("Windows блокирует файлы библиотек (Smart App Control): Параметры -> "
                     "Конфиденциальность и защита -> Безопасность Windows -> Управление "
                     "приложениями и браузером -> Интеллектуальное управление приложениями "
                     "-> выключить, затем запустить установку снова. Учтите: включить "
                     "эту защиту обратно можно только переустановкой Windows.")
    elif "DLL load failed" in everything or "0xC0000135" in everything.upper():
        hints.append("Не хватает системных библиотек Microsoft Visual C++. Установите "
                     "https://aka.ms/vs/17/release/vc_redist.x64.exe и запустите установку "
                     "снова.")
    if any(module == "tkinter" for module, _ in failed):
        hints.append("Модуль tkinter входит в Python: переустановите Python с "
                     "https://www.python.org/downloads/, оставив пункт «tcl/tk and IDLE».")
    if not hints:
        hints.append("Удалите папку .venv в папке приложения и запустите установку снова. "
                     "Если не поможет, пришлите текст ошибки выше.")
    fail("библиотеки установлены, но не загружаются (текст ошибки выше).\n" + "\n".join(hints))


def windows_shortcut():
    """Ярлык на рабочем столе. Пути передаются через переменные окружения."""
    script = (
        "$desktop = [Environment]::GetFolderPath('Desktop');"
        "$link = (New-Object -ComObject WScript.Shell).CreateShortcut("
        "(Join-Path $desktop ($env:SVERKA_NAME + '.lnk')));"
        "$link.TargetPath = $env:SVERKA_TARGET;"
        "$link.Arguments = '\"' + $env:SVERKA_SCRIPT + '\"';"
        "$link.WorkingDirectory = $env:SVERKA_DIR;"
        "$link.Description = $env:SVERKA_NAME;"
        "$link.Save();"
        "Write-Output $link.FullName")
    # Если ярлык с русским названием создать не удалось, пробуем латиницу
    for name in (APP_NAME, "Sverka Excel"):
        env = dict(os.environ, SVERKA_NAME=name, SVERKA_TARGET=str(venv_python(True)),
                   SVERKA_SCRIPT=str(GUI), SVERKA_DIR=str(APP_DIR))
        try:
            done = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
                env=env, capture_output=True, text=True)
        except OSError:
            break
        if done.returncode == 0 and Path(done.stdout.strip()).is_file():
            return f"ярлык «{name}» на рабочем столе"
    return windows_launcher_file()


def windows_launcher_file():
    """Запасной вариант, если ярлык создать не дали: файл запуска на рабочем столе."""
    import ctypes
    buffer = ctypes.create_unicode_buffer(260)
    if ctypes.windll.shell32.SHGetFolderPathW(None, 0x10, None, 0, buffer) != 0:
        return None
    target = Path(buffer.value) / f"{APP_NAME}.bat"
    target.write_text("@echo off\r\nchcp 65001 >nul\r\n"
                      f'start "" "{venv_python(True)}" "{GUI}"\r\n',
                      encoding="utf-8", newline="")
    return f"файл «{target.name}» на рабочем столе"


def macos_app():
    """Небольшое приложение-ярлык на рабочем столе (запускает окно без терминала)."""
    desktop = Path.home() / "Desktop"
    if not desktop.is_dir():
        return None
    app = desktop / f"{APP_NAME}.app"
    shutil.rmtree(app, ignore_errors=True)
    binary = app / "Contents" / "MacOS" / "sverka"
    binary.parent.mkdir(parents=True)
    binary.write_text(f'#!/bin/sh\ncd "{APP_DIR}"\nexec "{venv_python()}" "{GUI}"\n',
                      encoding="utf-8")
    binary.chmod(0o755)
    with open(app / "Contents" / "Info.plist", "wb") as f:
        plistlib.dump({"CFBundleName": APP_NAME, "CFBundleDisplayName": APP_NAME,
                       "CFBundleIdentifier": "local.sverka.excel",
                       "CFBundleExecutable": "sverka", "CFBundlePackageType": "APPL",
                       "CFBundleVersion": "1.0", "CFBundleShortVersionString": "1.0",
                       "NSHighResolutionCapable": True}, f)
    return f"приложение «{APP_NAME}» на рабочем столе"


def linux_desktop_file():
    entry = ("[Desktop Entry]\nType=Application\n"
             f"Name={APP_NAME}\nComment=Сверка данных в двух Excel-файлах\n"
             f'Exec="{venv_python()}" "{GUI}"\nPath={APP_DIR}\n'
             "Terminal=false\nCategories=Office;\n")
    menu = Path.home() / ".local" / "share" / "applications"
    menu.mkdir(parents=True, exist_ok=True)
    targets = [menu / "sverka-excel.desktop"]
    desktop = Path.home() / "Desktop"
    try:
        found = subprocess.run(["xdg-user-dir", "DESKTOP"], capture_output=True, text=True)
        if found.returncode == 0 and found.stdout.strip():
            desktop = Path(found.stdout.strip())
    except OSError:
        pass
    if desktop.is_dir() and desktop != Path.home():
        targets.append(desktop / "sverka-excel.desktop")
    for target in targets:
        target.write_text(entry, encoding="utf-8")
        target.chmod(target.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)
    return (f"пункт «{APP_NAME}» в меню приложений"
            + (" и ярлык на рабочем столе" if len(targets) > 1 else ""))


def create_launcher():
    try:
        if WINDOWS:
            return windows_shortcut()
        if MACOS:
            return macos_app()
        return linux_desktop_file()
    except OSError as e:
        print(f"Ярлык создать не удалось: {e}")
        return None


def main():
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass
    print(f"Установка приложения «{APP_NAME}»\nПапка: {APP_DIR}\n")
    check_python()
    create_venv()
    install_packages()
    launcher = create_launcher()
    fallback = "run.bat" if WINDOWS else "run.sh"
    print("\nГотово. Приложение установлено.")
    if launcher:
        print(f"Запуск: {launcher}.")
        print(f"Запасной способ: файл {fallback} в папке приложения.")
    else:
        print(f"Запуск: файл {fallback} в папке приложения.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
