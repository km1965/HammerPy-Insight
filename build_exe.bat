@echo off
REM ================================================================
REM  MK_Hammerpy_PRO_V01 - Build Script
REM  HammerPy Insight v3.0 — Phase 3 complète
REM  Génère un .exe autonome (onefile) avec icône
REM ================================================================

setlocal enabledelayedexpansion
chcp 65001 >nul

REM ── Couleurs console ──────────────────────────────────────────
echo.
echo ================================================================
echo   MK_Hammerpy_PRO_V01 — Build Script
echo   HammerPy Insight v3.0
echo ================================================================
echo.

REM ── Vérifications préliminaires ───────────────────────────────
echo [1/5] Vérification de Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo   ERREUR : Python non trouve dans le PATH.
    echo   Installez Python 3.10+ depuis https://python.org
    pause
    exit /b 1
)
python --version
echo.

echo [2/5] Vérification de PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo   PyInstaller non installe. Installation en cours...
    pip install pyinstaller
    if errorlevel 1 (
        echo   ERREUR : Echec de l'installation de PyInstaller.
        pause
        exit /b 1
    )
)
pyinstaller --version
echo.

REM ── Nettoyage des builds précédents ───────────────────────────
echo [3/5] Nettoyage des builds precedents...
if exist "build"      rmdir /s /q "build"
if exist "dist"       rmdir /s /q "dist"
if exist "MK_Hammerpy_PRO_V01.spec" del /q "MK_Hammerpy_PRO_V01.spec"
echo   Fichiers nettoyes.
echo.

REM ── Vérification de l'icône ───────────────────────────────────
echo [4/5] Verification de l'icone...
if not exist "hammerpy_icon.ico" (
    echo   ATTENTION : hammerpy_icon.ico introuvable.
    echo   Le .exe sera cree sans icone personnalisee.
    set ICON_ARG=
) else (
    echo   Icone trouvee : hammerpy_icon.ico
    set ICON_ARG=--icon=hammerpy_icon.ico
)
echo.

REM ── Construction du .exe ──────────────────────────────────────
echo [5/5] Construction de l'executable...
echo   Cela peut prendre 2 a 5 minutes...
echo.

pyinstaller ^
    --onefile ^
    --windowed ^
    --name "MK_Hammerpy_PRO_V01" ^
    %ICON_ARG% ^
    --add-data "exemple_profil_en_long.csv;." ^
    --add-data "air_valve_sizing.py;." ^
    --add-data "data_parser.py;." ^
    --add-data "pump_parser.py;." ^
    --add-data "report_generator.py;." ^
    --add-data "utils.py;." ^
    --add-data "workbook.py;." ^
    --hidden-import "customtkinter" ^
    --hidden-import "matplotlib" ^
    --hidden-import "matplotlib.backends.backend_tkagg" ^
    --hidden-import "pandas" ^
    --hidden-import "openpyxl" ^
    --hidden-import "docx" ^
    --hidden-import "PIL" ^
    --collect-all "customtkinter" ^
    --collect-all "matplotlib" ^
    --noconfirm ^
    main.py

if errorlevel 1 (
    echo.
    echo   ERREUR : Echec de la construction de l'executable.
    pause
    exit /b 1
)

echo.
echo ================================================================
echo   BUILD TERMINE AVEC SUCCES
echo ================================================================
echo.
echo   Executable : dist\MK_Hammerpy_PRO_V01.exe
echo.

REM ── Affichage de la taille du fichier ────────────────────────
if exist "dist\MK_Hammerpy_PRO_V01.exe" (
    for %%I in ("dist\MK_Hammerpy_PRO_V01.exe") do (
        echo   Taille : %%~zI octets ^(~%%~zI / 1048576 Mo^)
    )
)

echo.
echo   Vous pouvez maintenant distribuer dist\MK_Hammerpy_PRO_V01.exe
echo.
pause
endlocal
