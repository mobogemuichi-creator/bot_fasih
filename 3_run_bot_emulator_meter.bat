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
echo [1/2] Memeriksa & mengunduh pembaruan dari GitHub...
git pull origin main
echo.
echo [2/2] Memulai bot emulator meter...
echo.

python bot_emulator_meter.py

echo.
echo ========================================================
echo Script telah selesai atau terhenti.
echo Tekan tombol apa saja untuk menutup jendela ini...
echo ========================================================
pause
