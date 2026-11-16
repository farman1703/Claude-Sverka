@echo off
setlocal
cd /d "%~dp0"

set "PY="
py -3 -c "import sys; sys.exit(sys.version_info < (3, 9))" >nul 2>nul && set "PY=py -3"
if not defined PY python -c "import sys; sys.exit(sys.version_info < (3, 9))" >nul 2>nul && set "PY=python"
if not defined PY goto nopython

%PY% install.py
if errorlevel 1 goto failed
echo.
pause
exit /b 0

:nopython
echo Python 3.9 or newer was not found on this computer.
echo 1. Download and install Python: https://www.python.org/downloads/
echo    In the installer tick "Add python.exe to PATH".
echo 2. Run install.bat again.
echo.
pause
exit /b 1

:failed
echo.
echo Installation failed. See the message above.
pause
exit /b 1
