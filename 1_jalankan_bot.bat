@echo off
title Bot FASIH - Auto Approve
cd /d "%~dp0"

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python tidak terdeteksi di PATH.
    pause
    exit /b
)

python bot_fasih.py

echo.
echo Proses selesai atau dihentikan.
pause
