@echo off
title Setup Dependensi Bot FASIH
cd /d "%~dp0"

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python tidak terdeteksi. Install Python dari https://www.python.org/downloads/
    echo dan centang 'Add python.exe to PATH'.
    pause
    exit /b
)

echo [1/3] Install library Python...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Gagal install requirements.txt
    pause
    exit /b
)

echo [2/3] Install browser Playwright Chromium...
python -m playwright install chromium

echo [3/3] Cek file akun SSO...
if not exist sso (
    if exist sso.example (
        copy sso.example sso >nul
        echo File 'sso' dibuat dari sso.example. Isi username dan password SSO BPS, lalu simpan.
        start notepad sso
    )
) else (
    echo File 'sso' sudah ada.
)

echo.
echo Selesai. Taruh file Excel sebagai 'daftar.xlsx' di folder ini, lalu jalankan 1_jalankan_bot.bat
pause
