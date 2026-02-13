@echo off
TITLE PGT-A Report Generator

cd /d "%~dp0"

echo =========================================
echo   PGT-A Report Generator
echo =========================================
echo.

where python >nul 2>&1
if %errorlevel% neq 0 (
    where py >nul 2>&1
    if %errorlevel% neq 0 (
        echo ERROR: Python not found
        echo Install from https://www.python.org
        pause
        exit
    )
)

echo Python found.
echo.

if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
    if %errorlevel% neq 0 (
        py -m venv .venv
    )
)

echo Activating environment...
call .venv\Scripts\activate.bat

python -c "import PyQt6" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing core packages...
    pip install --upgrade pip
    pip install PyQt6>=6.6.0 PyQt6-Qt6>=6.6.0 reportlab>=4.0.0 PyPDF2>=3.0.0 pdfplumber>=0.10.0 python-docx>=1.0.0 pandas>=2.0.0 openpyxl>=3.1.0 pyarrow>=14.0.0 Pillow>=10.0.0 numpy>=1.24.0 requests>=2.31.0
    echo.
    echo Core packages installed. Starting app...
    echo EasyOCR will be installed in background later.
)

echo.
echo Starting application...
echo.

python pgta_report_generator.py

echo.
echo =========================================
echo Application closed.
echo =========================================
pause
