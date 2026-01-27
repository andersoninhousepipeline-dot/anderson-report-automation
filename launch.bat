@echo off
TITLE PGT-A Report Generator
echo ===================================
echo PGT-A Report Generator (Windows)
echo ===================================
echo.

:: Check for Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not in your PATH.
    echo Please install Python 3.10+ from python.org and check "Add to PATH".
    pause
    exit /b
)

:: Check dependencies
echo Checking dependencies...
python -c "import reportlab, docx, PyQt6, pandas, openpyxl, PIL, pdfplumber, PyPDF2" >nul 2>&1
if %errorlevel% neq 0 (
    echo Missing dependencies detected!
    echo Installing required packages...
    pip install -r requirements_app.txt
    if %errorlevel% neq 0 (
        echo Error: Failed to install dependencies.
        pause
        exit /b
    )
)

echo.
echo Starting PGT-A Report Generator...
echo.
python pgta_report_generator.py

echo.
echo Application closed.
pause
