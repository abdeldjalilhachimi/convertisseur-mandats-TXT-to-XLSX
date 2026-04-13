@echo off
setlocal
title Build Mandats TXT-to-XLSX

echo [1/3] Installing dependencies...
pip install -r requirements.txt pyinstaller --quiet
if errorlevel 1 ( echo ERROR: pip failed & pause & exit /b 1 )

echo [2/3] Building executable...
pyinstaller app.spec --clean --noconfirm
if errorlevel 1 ( echo ERROR: PyInstaller failed & pause & exit /b 1 )

echo [3/3] Done!
echo.
echo Output: dist\Mandats-TXT-to-XLSX.exe
echo.
pause
