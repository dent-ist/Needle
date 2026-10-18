@echo off
chcp 65001 > nul
cd /d "%~dp0"

rem --- 사용할 Python 찾기 (py 런처 우선) ---
set "PY="
py -3 -c "import sys" >nul 2>&1 && set "PY=py -3"
if not defined PY python -c "import sys" >nul 2>&1 && set "PY=python"
if not defined PY (
  echo Python을 찾을 수 없습니다. python.org 에서 설치한 뒤 다시 실행해 주세요.
  goto fail
)

rem --- 빌드 전용 가상환경 (.venv) ---
if not exist ".venv\Scripts\python.exe" (
  echo [1/3] 빌드용 가상환경 만드는 중...
  %PY% -m venv .venv
  if errorlevel 1 goto fail
)
set "VPY=%~dp0.venv\Scripts\python.exe"

echo [2/3] 필요한 패키지 설치 중...
"%VPY%" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto fail

echo [3/3] 실행 파일 만드는 중... (1~3분)
"%VPY%" -m PyInstaller --noconfirm --onefile --windowed --name Needle --icon needle.ico --hidden-import uuid --hidden-import decimal --hidden-import datetime --hidden-import zoneinfo --hidden-import ipaddress --collect-submodules duckdb --exclude-module numpy --exclude-module pandas --exclude-module pyarrow --exclude-module matplotlib --exclude-module scipy --exclude-module IPython --exclude-module PIL --exclude-module polars csv_search.py
if errorlevel 1 goto fail

echo.
echo 완료: %~dp0dist\Needle.exe
pause
exit /b 0

:fail
echo.
echo 빌드에 실패했습니다. 위 메시지를 확인해 주세요.
pause
exit /b 1
