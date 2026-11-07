@echo off
TITLE He Thong Lam Form DAS V3.0
CHCP 65001 > nul
cd /d "%~dp0"

echo ======================================================
echo   DANG KHOI DONG HE THONG LAM FORM (DAS V3.0)
echo ======================================================

set "PY_PATH="
set "PY_DIR="

if exist "%LOCALAPPDATA%\Python\bin\python.exe" (
    set "PY_PATH=%LOCALAPPDATA%\Python\bin\python.exe"
    set "PY_DIR=%LOCALAPPDATA%\Python\bin"
    goto :found_python
)
if exist "%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe" (
    set "PY_PATH=%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe"
    set "PY_DIR=%LOCALAPPDATA%\Python\pythoncore-3.14-64"
    goto :found_python
)
if exist "C:\Users\rik_7\AppData\Local\Python\bin\python.exe" (
    set "PY_PATH=C:\Users\rik_7\AppData\Local\Python\bin\python.exe"
    set "PY_DIR=C:\Users\rik_7\AppData\Local\Python\bin"
    goto :found_python
)
if exist "C:\Users\rik_7\AppData\Local\Python\pythoncore-3.14-64\python.exe" (
    set "PY_PATH=C:\Users\rik_7\AppData\Local\Python\pythoncore-3.14-64\python.exe"
    set "PY_DIR=C:\Users\rik_7\AppData\Local\Python\pythoncore-3.14-64"
    goto :found_python
)

set "PY_PATH=python"

:found_python
if defined PY_DIR (
    set "PATH=%PY_DIR%;%PY_DIR%\Scripts;%PATH%"
)

echo [1/3] Kiem tra Python...
"%PY_PATH%" --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo LOI: Khong the khoi chay Python!
    echo Vui long kiem tra lai moi truong Python tren may.
    echo.
    pause
    exit /b
)

for /f "tokens=*" %%V in ('"%PY_PATH%" --version') do echo [OK] %%V

echo [2/3] Kiem tra thu vien he thong...
"%PY_PATH%" -c "import flask, flask_cors, deep_translator, sqlalchemy, PIL, openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo Dang cai dat thu vien can thiet...
    "%PY_PATH%" -m pip install -r requirements.txt
) else (
    echo [OK] Da co day du cac thu vien can thiet.
)

echo [3/3] Dang khoi dong Server tai http://127.0.0.1:5000 ...
start http://127.0.0.1:5000/demo-das-form
"%PY_PATH%" app.py

pause
