@echo off
TITLE PGT-A Report Generator
cd /d "%~dp0"

echo.
echo ===========================================
echo   PGT-A Report Generator (Windows)
echo ===========================================
echo.

REM Check for Python
python --version >nul 2>&1
if errorlevel 1 (
    py --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python is not installed!
        echo.
        echo Please install Python from https://www.python.org/downloads/
        echo Make sure to check "Add Python to PATH" during installation.
        echo.
        goto :end
    )
    set PYCMD=py
) else (
    set PYCMD=python
)

echo [OK] Python found
%PYCMD% --version
echo.

REM Create virtual environment if needed
if not exist ".venv\Scripts\python.exe" (
    echo [*] First time setup - creating virtual environment...
    %PYCMD% -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment
        goto :end
    )
    echo [OK] Created
    echo.
)

REM Activate virtual environment
echo [*] Activating environment...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo [ERROR] Failed to activate environment
    goto :end
)
echo [OK] Activated
echo.

REM Check if core dependencies need to be installed
python -c "import PyQt6" >nul 2>&1
if errorlevel 1 (
    echo [*] Installing dependencies (first run)...
    echo     This may take several minutes.
    echo.
    python -m pip install --upgrade pip
    python -m pip install PyQt6 reportlab PyPDF2 pdfplumber python-docx pandas openpyxl pyarrow Pillow numpy requests
    if errorlevel 1 (
        echo.
        echo [ERROR] Failed to install core dependencies
        goto :end
    )
    echo.
    echo [*] Installing EasyOCR (optional)...
    python -m pip install easyocr
    echo.
    echo [OK] Dependencies installed
    echo.
)

echo ===========================================
echo [*] Starting Application...
echo ===========================================
echo.

REM Run the application
python pgta_report_generator.py

echo.
echo ===========================================
echo Application closed.
echo ===========================================

:end
echo.
echo Press any key to exit...
pause >nul
