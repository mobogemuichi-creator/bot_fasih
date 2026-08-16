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
    echo [INFO] Mengunduh dan meng-install Git secara otomatis via Windows Winget...
    echo.
    winget install --id Git.Git -e --source winget --accept-source-agreements --accept-package-agreements
    echo.
    echo ========================================================
    echo Git telah berhasil di-install!
    echo Silakan TUTUP jendela ini dan JALANKAN ULANG file .bat ini.
    echo ========================================================
    pause
    exit /b
) else (
    set "HAS_KONFIG="
    if exist "konfigurasi.py" (
        set "HAS_KONFIG=1"
        copy /y "konfigurasi.py" "konfigurasi_backup.tmp" >nul 2>&1
    )

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
        echo [1/2] Memeriksa dan mengunduh pembaruan dari GitHub...
        git fetch origin main
        git reset --hard origin/main
        echo.
    )

    if defined HAS_KONFIG (
        if exist "konfigurasi_backup.tmp" (
            copy /y "konfigurasi_backup.tmp" "konfigurasi.py" >nul 2>&1
            del /f /q "konfigurasi_backup.tmp" >nul 2>&1
        )
    )

    if exist "konfigurasi.py" (
        git update-index --skip-worktree konfigurasi.py >nul 2>&1
    )
)

echo [2/2] Memulai bot emulator...
echo.

python bot_emulator_reject_list.py

echo.
echo ========================================================
echo Script telah selesai atau terhenti.
echo Tekan tombol apa saja untuk menutup jendela ini...
echo ========================================================
pause
