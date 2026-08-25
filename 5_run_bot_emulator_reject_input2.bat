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

title Bot Emulator Reject Input 2 Runner
echo ========================================================
echo      MENJALANKAN BOT EMULATOR REJECT INPUT 2
echo ========================================================
echo.

:: Pembaruan proyek & git terpusat lewat auto_update.bat
call "%~dp0auto_update.bat" --no-pause

echo [2/2] Memulai bot emulator reject input 2...
echo.

python bot_emulator_reject_input2.py

echo.
echo ========================================================
echo Script telah selesai atau terhenti.
echo Tekan tombol apa saja untuk menutup jendela ini...
echo ========================================================
pause
