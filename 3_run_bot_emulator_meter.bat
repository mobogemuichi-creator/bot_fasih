@echo off
:: Memastikan script berjalan sebagai Administrator (UAC prompt otomatis)
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Meminta izin Administrator...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Pindah ke folder script berada
cd /d "%~dp0"

title Bot Emulator Meter Runner
echo ========================================================
echo            MENJALANKAN BOT EMULATOR METER
echo ========================================================
echo.

where git >nul 2>&1
if %errorLevel% neq 0 (
    echo [PERINGATAN] Git belum ter-install di komputer ini.
    echo Harap install Git dari https://git-scm.com/ agar bisa auto-update.
    echo.
) else (
    if not exist ".git" (
        echo [1/2] Pertama kali dijalankan! Mengunduh seluruh project dari GitHub...
        git init
        git remote add origin https://github.com/mobogemuichi-creator/bot_fasih.git
        git fetch origin main
        git reset --hard origin/main
        git branch -M main
        git branch --set-upstream-to=origin/main main
        echo.
    ) else (
        echo [1/2] Memeriksa & mengunduh pembaruan dari GitHub...
        git pull origin main
        echo.
    )
)

echo [2/2] Memulai bot emulator meter...
echo.

python bot_emulator_meter.py

echo.
echo ========================================================
echo Script telah selesai atau terhenti.
echo Tekan tombol apa saja untuk menutup jendela ini...
echo ========================================================
pause
