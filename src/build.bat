@echo off
chcp 65001 > nul
cd /d "%~dp0"

rem --- Search Python (py launcher first) ---
set "PY="
py -3 -c "import sys" >nul 2>&1 && set "PY=py -3"
if not defined PY python -c "import sys" >nul 2>&1 && set "PY=python"
if not defined PY (
  echo Cannot find Python. Please install from python.org and re-run
  goto fail
)

rem --- Build-only virtual environment (.venv) ---
if not exist ".venv\Scripts\python.exe" (
  echo [1/3] Creating build virtual environment
  %PY% -m venv .venv
  if errorlevel 1 goto fail
)
set "VPY=%~dp0.venv\Scripts\python.exe"

echo [2/3] Installing required packages...
"%VPY%" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto fail

echo [3/3] Creating executable file (1~3 minutes)
"%VPY%" -m PyInstaller --noconfirm --onefile --windowed --name CsvFastSearch --hidden-import uuid --hidden-import decimal --hidden-import datetime --hidden-import zoneinfo --hidden-import ipaddress --collect-submodules duckdb --exclude-module numpy --exclude-module pandas --exclude-module pyarrow --exclude-module matplotlib --exclude-module scipy --exclude-module IPython --exclude-module PIL --exclude-module polars csv_search.py
if errorlevel 1 goto fail

echo.
echo Completed: %~dp0dist\CsvFastSearch.exe
pause
exit /b 0

:fail
echo.
echo Build failed. please see message above
pause
exit /b 1
