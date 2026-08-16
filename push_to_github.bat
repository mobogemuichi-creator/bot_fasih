@echo off
cd /d "%~dp0"
title Push to GitHub Runner
echo ========================================================
echo               PUSH PERUBAHAN KE GITHUB
echo ========================================================
echo.

python push_to_github.py

echo.
echo ========================================================
echo Selesai. Tekan tombol apa saja untuk menutup...
echo ========================================================
pause
