#!/bin/sh
# Запуск приложения «Сверка Excel» (macOS, Linux)
cd "$(dirname "$0")" || exit 1
if [ ! -x .venv/bin/python ]; then
    echo "Приложение ещё не установлено. Сначала запустите install.sh (macOS: install.command)."
    exit 1
fi
exec .venv/bin/python sverka_gui.py
