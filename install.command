#!/bin/sh
# Установка приложения «Сверка Excel» (macOS, Linux)
cd "$(dirname "$0")" || exit 1
if command -v python3 >/dev/null 2>&1; then
    python3 install.py
    status=$?
else
    echo "Python 3 не найден. Установите Python 3.9 или новее:"
    echo "  macOS: https://www.python.org/downloads/"
    echo "  Linux: sudo apt install python3 python3-venv python3-tk"
    status=1
fi
echo
printf "Нажмите Enter, чтобы закрыть окно... "
read -r _
exit $status
