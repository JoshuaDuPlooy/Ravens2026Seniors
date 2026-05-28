@echo off
cd /d "%~dp0"
python -c "import openpyxl" 2>nul || (
    echo Installing required package: openpyxl...
    pip install openpyxl
)
python ExcelToJSON.py
