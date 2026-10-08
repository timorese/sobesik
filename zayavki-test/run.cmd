@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
    where python >nul 2>nul
    if errorlevel 1 (
        echo Python not found. Install Python 3.10+ from python.org and enable PATH.
        pause
        exit /b 1
    )
    set "ZAYAVKI_PYTHON=python"
) else (
    set "ZAYAVKI_PYTHON=py -3"
)
%ZAYAVKI_PYTHON% -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)"
if errorlevel 1 goto :python_error
if not exist ".venv\Scripts\python.exe" (
    %ZAYAVKI_PYTHON% -m venv .venv
    if errorlevel 1 goto :failed
)
".venv\Scripts\python.exe" -c "import openpyxl; assert openpyxl.__version__ == '3.1.5'" >nul 2>nul
if errorlevel 1 (
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 goto :failed
)
if "%~1"=="" (
    ".venv\Scripts\python.exe" clean_zayavki.py zayavki.xlsx --period 2026-09 --out-dir results --force
) else (
    ".venv\Scripts\python.exe" clean_zayavki.py %*
)
if errorlevel 1 goto :failed
echo Done. See output paths above.
pause
exit /b 0
:python_error
echo Python 3.10 or newer is required.
pause
exit /b 1
:failed
echo Failed. Read the error above. Close Excel before overwriting output files.
pause
exit /b 1
