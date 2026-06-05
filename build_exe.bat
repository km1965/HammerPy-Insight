@echo off
REM ============================================
REM  Build MK_hammerpy_PRO_V01.exe avec PyInstaller
REM ============================================
echo.
echo === Nettoyage des builds precedents ===
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build

echo.
echo === Installation des dependances ===
pip install -r requirements.txt

echo.
echo === Generation de l'executable ===
pyinstaller --clean --onefile --windowed --icon=hammerpy_icon.ico --name=MK_hammerpy_PRO_V01 --exclude-module PyQt5 --exclude-module PyQt6 --exclude-module PySide2 --exclude-module PySide6 main.py

echo.
echo === Nettoyage des fichiers temporaires ===
if exist "build" rmdir /s /q build

echo.
echo === Termine ! ===
echo L'executable se trouve dans : dist\MK_hammerpy_PRO_V01.exe
pause
