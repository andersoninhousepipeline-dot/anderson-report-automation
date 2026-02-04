@echo off
SETLOCAL EnableDelayedExpansion
TITLE PGT-A Report Generator - Setup
chcp 65001 >nul 2>&1

echo.
echo ===========================================
echo   PGT-A Report Generator - SETUP
echo ===========================================
echo.
echo This will set up the application for first use.
echo.

:: Change to script directory
cd /d "%~dp0"
echo Working directory: %CD%
echo.

:: ===== CHECK FOR PYTHON =====
echo [*] Checking for Python installation...

set "PYTHON_CMD="

where python >nul 2>&1
if %errorlevel% equ 0 (
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PYVER=%%i
    echo     Found: !PYVER!
    set "PYTHON_CMD=python"
    goto :python_ok
)

where py >nul 2>&1
if %errorlevel% equ 0 (
    for /f "tokens=*" %%i in ('py --version 2^>^&1') do set PYVER=%%i
    echo     Found: !PYVER!
    set "PYTHON_CMD=py"
    goto :python_ok
)

echo.
echo [X] ERROR: Python is not installed!
echo.
echo Please install Python first:
echo   1. Go to https://www.python.org/downloads/
echo   2. Download Python 3.10 or newer (64-bit recommended)
echo   3. Run the installer
echo   4. IMPORTANT: Check "Add Python to PATH" at the bottom!
echo   5. Click "Install Now"
echo   6. Restart your computer
echo   7. Run this setup again
echo.
echo Press any key to open the Python download page...
pause >nul
start https://www.python.org/downloads/
exit /b 1

:python_ok
echo [OK] Python is installed
echo.

:: ===== CHECK ARCHITECTURE =====
%PYTHON_CMD% -c "import struct; bits=struct.calcsize('P')*8; print(f'[OK] {bits}-bit Python')"
echo.

:: ===== REMOVE OLD VENV IF EXISTS =====
if exist ".venv" (
    echo [*] Removing old virtual environment...
    rmdir /s /q .venv 2>nul
    if exist ".venv" (
        echo [!] Could not fully remove .venv folder
        echo     Please close any programs using it and try again
        pause
        exit /b 1
    )
    echo [OK] Old environment removed
    echo.
)

:: ===== CREATE NEW VIRTUAL ENVIRONMENT =====
echo [*] Creating new virtual environment...
%PYTHON_CMD% -m venv .venv
if !errorlevel! neq 0 (
    echo.
    echo [X] ERROR: Failed to create virtual environment!
    echo.
    echo Try these solutions:
    echo   1. Run as Administrator
    echo   2. Make sure Python was installed with "pip" option
    echo   3. Try: %PYTHON_CMD% -m pip install --user virtualenv
    echo.
    pause
    exit /b 1
)
echo [OK] Virtual environment created
echo.

:: ===== ACTIVATE ENVIRONMENT =====
echo [*] Activating environment...
call .venv\Scripts\activate.bat
echo [OK] Activated
echo.

:: ===== UPGRADE PIP =====
echo [*] Upgrading pip...
python -m pip install --upgrade pip --quiet
echo [OK] Pip upgraded
echo.

:: ===== INSTALL DEPENDENCIES =====
echo [*] Installing dependencies...
echo.
echo     This will download and install:
echo       - PyQt6 (GUI framework)
echo       - EasyOCR (for TRF verification)
echo       - ReportLab (PDF generation)
echo       - And other required packages
echo.
echo     This may take 5-15 minutes depending on your internet speed.
echo     Please wait...
echo.

python -m pip install -r requirements.txt

if !errorlevel! neq 0 (
    echo.
    echo [X] ERROR: Failed to install some dependencies!
    echo.
    echo Possible solutions:
    echo   1. Check your internet connection
    echo   2. Run this script as Administrator
    echo   3. Try installing manually:
    echo      .venv\Scripts\activate
    echo      pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

echo.
echo [OK] All dependencies installed!
echo.

:: ===== VERIFY INSTALLATION =====
echo [*] Verifying installation...
python -c "import PyQt6; print('    PyQt6:', PyQt6.QtCore.PYQT_VERSION_STR)"
python -c "import reportlab; print('    ReportLab:', reportlab.Version)"
python -c "import pandas; print('    Pandas:', pandas.__version__)"
python -c "import PIL; print('    Pillow:', PIL.__version__)"
python -c "import easyocr; print('    EasyOCR: installed')" 2>nul || echo     EasyOCR: will download models on first use

echo.
echo ===========================================
echo   SETUP COMPLETE!
echo ===========================================
echo.
echo To run the application:
echo   - Double-click "launch.bat"
echo.
echo Or from Command Prompt:
echo   - cd "%CD%"
echo   - .venv\Scripts\activate
echo   - python pgta_report_generator.py
echo.

call .venv\Scripts\deactivate.bat 2>nul

echo Press any key to exit...
pause >nul
ENDLOCAL

