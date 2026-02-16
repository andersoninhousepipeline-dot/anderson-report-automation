@echo off
TITLE PGT-A Report Generator

cd /d "%~dp0"

echo =========================================
echo   PGT-A Report Generator
echo =========================================
echo.

REM Check for Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    where py >nul 2>&1
    if %errorlevel% neq 0 (
        echo [ERROR] Python not found
        echo Please install from https://www.python.org/downloads/
        pause
        exit
    )
)

echo Python found.
echo.

REM Create venv if missing
if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
    if %errorlevel% neq 0 (
        py -m venv .venv
    )
)

echo Activating environment...
call .venv\Scripts\activate.bat

REM Architecture Check (Crucial for PyQt6)
python -c "import struct; import sys; sys.exit(0 if (struct.calcsize('P') * 8) == 64 else 1)" >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] 32-bit Python detected!
    echo PGT-A Report Generator requires 64-bit Python.
    echo Please uninstall Python and install the 64-bit version.
    echo.
    pause
    exit
)

REM Dependency Check & Install
python -c "import PyQt6" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing core packages...
    python -m pip install --upgrade pip
    python -m pip install PyQt6>=6.6.0 PyQt6-Qt6>=6.6.0 reportlab>=4.0.0 PyPDF2>=3.0.0 pdfplumber>=0.10.0 python-docx>=1.0.0 pandas>=2.0.0 openpyxl>=3.1.0 Pillow>=10.0.0 numpy>=1.24.0 requests>=2.31.0
    echo.
    echo Core packages installed. Starting app...
    echo EasyOCR will be installed in background later if enabled.
)

echo.
echo Starting application...
echo.

REM Launch with stderr capture
python pgta_report_generator.py 2> run_error.log

if %errorlevel% neq 0 (
    echo.
    echo =========================================
    echo   [ERROR] Application crashed!
    echo =========================================
    echo.
    echo Error details:
    type run_error.log
    echo.
    echo.
    pause
    exit
)

echo.
echo =========================================
echo Application closed.
echo =========================================
pause
