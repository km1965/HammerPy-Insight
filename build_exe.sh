#!/bin/bash
# Build MK_hammerpy_PRO_V01 avec PyInstaller
echo "=== Cleaning previous builds ==="
rm -rf dist build MK_hammerpy_PRO_V01.spec 2>/dev/null

echo ""
echo "=== Installing dependencies ==="
pip install -r requirements.txt

echo ""
echo "=== Generating executable ==="
pyinstaller --onefile --windowed \
    --icon="hammerpy_icon.ico" \
    --name "MK_hammerpy_PRO_V01" \
    --exclude-module PyQt5 \
    --exclude-module PyQt6 \
    --exclude-module PySide2 \
    --exclude-module PySide6 \
    main.py

echo ""
echo "=== Cleaning temp files ==="
rm -rf build

echo ""
echo "=== Done! ==="
echo "Executable: dist/MK_hammerpy_PRO_V01"
