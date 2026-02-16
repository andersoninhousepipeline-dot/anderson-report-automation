@echo off
TITLE PGT-A Report Generator - Setup
cd /d "%~dp0"

echo.
echo ===========================================
echo   PGT-A Report Generator - SETUP
echo ===========================================
echo.

REM Check for Python
python --version >nul 2>&1
if errorlevel 1 (
    py --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python is not installed!
        echo.
        echo Please:
        echo   1. Go to https://www.python.org/downloads/
        echo   2. Download and install Python 3.10+
        echo   3. CHECK "Add Python to PATH" during installation
        echo   4. Restart your computer
        echo   5. Run this setup again
        echo.
        pause
        exit /b 1
    )
    set PYCMD=py
) else (
    set PYCMD=python
)

echo [OK] Python found
%PYCMD% --version
echo.

REM Check 64-bit architecture
echo [*] Checking Python architecture...
%PYCMD% -c "import struct; bits=struct.calcsize('P')*8; print(f'  Architecture: {bits}-bit'); exit(0 if bits==64 else 1)"
if errorlevel 1 (
    echo [ERROR] 32-bit Python detected! PyQt6 requires 64-bit Python.
    echo         Please uninstall Python and reinstall the 64-bit version.
    echo         Download: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo.

REM Remove old venv if exists
if exist ".venv" (
    echo [*] Removing old virtual environment...
    rmdir /s /q .venv 2>nul
    timeout /t 2 >nul
    echo [OK] Removed
    echo.
)

REM Create new venv
echo [*] Creating virtual environment...
%PYCMD% -m venv .venv
if errorlevel 1 (
    echo [ERROR] Failed to create virtual environment
    pause
    exit /b 1
)
echo [OK] Created
echo.

REM Activate
echo [*] Activating environment...
call .venv\Scripts\activate.bat
echo [OK] Activated
echo.

REM Upgrade pip
echo [*] Upgrading pip...
python -m pip install --upgrade pip
echo.

REM Install core dependencies first (without easyocr)
echo [*] Installing core dependencies...
echo.
python -m pip install PyQt6 reportlab PyPDF2 pdfplumber python-docx pandas openpyxl Pillow numpy requests

if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install core dependencies
    pause
    exit /b 1
)
echo [OK] Core packages installed
echo.

REM Try to install EasyOCR (optional, may fail on some systems)
echo [*] Installing EasyOCR (optional, for TRF verification)...
echo     This may take a few minutes...
echo.
python -m pip install easyocr

if errorlevel 1 (
    echo.
    echo [!] EasyOCR installation had issues.
    echo     TRF verification will be disabled, but the app will work.
    echo.
    echo     To enable TRF verification later:
    echo       1. Install Visual C++ Redistributable:
    echo          https://aka.ms/vs/17/release/vc_redist.x64.exe
    echo       2. Delete .venv folder
    echo       3. Run setup.bat again
    echo.
) else (
    echo [OK] EasyOCR installed
)

echo.
echo ===========================================
echo   SETUP COMPLETE!
echo ===========================================
echo.
echo To run the application, double-click: launch.bat
echo.
echo NOTE: If you see PyTorch/DLL errors when running:
echo   Install Visual C++ Redistributable from:
echo   https://aka.ms/vs/17/release/vc_redist.x64.exe
echo.
pause


