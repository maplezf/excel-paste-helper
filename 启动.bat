@echo off
rem Pure ASCII on purpose: cmd parses .bat in the system codepage (GBK on
rem Chinese Windows). Any non-ASCII byte or "chcp 65001" mid-file makes the
rem parser misread the rest of the script and fail with garbled errors.
cd /d "%~dp0"
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "%~dp0main.py"
    exit /b 0
)
if exist ".venv\Scripts\python.exe" (
    start "" ".venv\Scripts\python.exe" "%~dp0main.py"
    exit /b 0
)
echo [ERROR] .venv not found in this folder.
echo Please open a terminal here and run:
echo     py -m venv .venv
echo     .venv\Scripts\pip install -r requirements.txt
pause
