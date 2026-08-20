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

title Auto Update Proyek
echo ========================================================
echo             OTOMATISASI UPDATE PROYEK DARI GITHUB
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
    echo Silakan TUTUP jendela ini dan JALANKAN ULANG file auto_update.bat ini.
    echo ========================================================
    pause
    exit /b
) else (
    if not exist ".git" (
        echo [1/2] Pertama kali dijalankan! Mengunduh seluruh project dari GitHub...
        git init
        git remote add origin https://github.com/mobogemuichi-creator/bot_fasih.git
        git fetch origin main
        git reset --hard origin/main
        git branch -M main
        git branch --set-upstream-to=origin/main main
        if exist "konfigurasi.py" (
            git update-index --skip-worktree konfigurasi.py >nul 2>&1
        )
        echo.
    ) else (
        echo [1/2] Memeriksa dan mengunduh pembaruan dari GitHub...
        if not exist "konfigurasi_lokal.py" if exist "konfigurasi.py" (
            echo [MIGRASI] Mengamankan konfigurasi lokal laptop ini ke konfigurasi_lokal.py...
            python -c "import konfigurasi as c; f=open('konfigurasi_lokal.py','w',encoding='utf-8'); f.write(f'# KONFIGURASI LOKAL LAPTOP INI\nLDPLAYER_DNCONSOLE = r\"{c.LDPLAYER_DNCONSOLE}\"\nLDPLAYER_ADB = r\"{c.LDPLAYER_ADB}\"\nFOTO_DIRECTORY = r\"{c.FOTO_DIRECTORY}\"\nFOTO_DIR = FOTO_DIRECTORY\nSLEEP_SHORT = {c.SLEEP_SHORT}\nSLEEP_MEDIUM = {c.SLEEP_MEDIUM}\nSLEEP_LONG = {c.SLEEP_LONG}\nSLEEP_LONG_REJECT = {c.SLEEP_LONG_REJECT}\n'); f.close()" >nul 2>&1
        )
        if exist "konfigurasi.py" (
            git update-index --no-skip-worktree konfigurasi.py >nul 2>&1
        )
        git fetch origin main
        git reset --hard origin/main
        if exist "konfigurasi.py" (
            git update-index --skip-worktree konfigurasi.py >nul 2>&1
        )
        echo.
    )
)

echo ========================================================
echo [SUKSES] Seluruh file proyek berhasil diperbarui!
if "%1"=="--no-pause" goto end_update
echo Tekan tombol apa saja untuk menutup jendela ini...
echo ========================================================
pause
:end_update
