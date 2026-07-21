@echo off
TITLE He Thong Lam Form DAS V3.0
CHCP 65001 > nul
cd /d "%~dp0"

echo ======================================================
echo   ĐANG KHỞI ĐỘNG HỆ THỐNG LÀM FORM (DAS V3.0)
echo ======================================================

set PY_PATH=python
if exist "C:\Users\rik_7\AppData\Local\Python\bin\python.exe" (
    set PY_PATH="C:\Users\rik_7\AppData\Local\Python\bin\python.exe"
)

echo [1/3] Kiểm tra Python...
%PY_PATH% --version >nul 2>&1
if %errorlevel% neq 0 (
    echo LỖI: Python chưa được cài đặt hoặc chưa được thêm vào PATH!
    pause
    exit /b
)

echo [2/3] Cài đặt/Cập nhật thư viện cần thiết...
%PY_PATH% -m pip install -r requirements.txt

echo [3/3] Đang chạy Server tại http://127.0.0.1:5000
echo Vui lòng đợi trong giây lát để hệ thống khởi động...
timeout /t 5 /nobreak > nul
start http://127.0.0.1:5000
%PY_PATH% app.py

pause