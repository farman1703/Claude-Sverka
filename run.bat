@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\pythonw.exe" goto notinstalled
start "" ".venv\Scripts\pythonw.exe" "%~dp0sverka_gui.py"
exit /b 0

:notinstalled
echo The application is not installed yet. Run install.bat first.
pause
exit /b 1
