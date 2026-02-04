@echo off
SETLOCAL EnableDelayedExpansion
TITLE PGT-A Report Generator
chcp 65001 >nul 2>&1

echo.
echo ===========================================
echo   PGT-A Report Generator (Windows)
echo ===========================================
echo.

:: Change to script directory (where this batch file is located)
cd /d "%~dp0"
echo Working directory: %CD%
echo.

:: ===== CHECK FOR PYTHON =====
echo [*] Checking for Python...

:: Try different Python commands
set "PYTHON_CMD="

where python >nul 2>&1
if %errorlevel% equ 0 (
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PYVER=%%i
    echo     Found: !PYVER!
    set "PYTHON_CMD=python"
    goto :python_found
)

where python3 >nul 2>&1
if %errorlevel% equ 0 (
    for /f "tokens=*" %%i in ('python3 --version 2^>^&1') do set PYVER=%%i
    echo     Found: !PYVER!
    set "PYTHON_CMD=python3"
    goto :python_found
)

where py >nul 2>&1
if %errorlevel% equ 0 (
    for /f "tokens=*" %%i in ('py --version 2^>^&1') do set PYVER=%%i
    echo     Found: !PYVER!
    set "PYTHON_CMD=py"
    goto :python_found
)

:: Python not found
echo.
echo [X] ERROR: Python is not installed or not in PATH!
echo.
echo Please install Python:
echo   1. Download from https://www.python.org/downloads/
echo   2. During installation, CHECK the box "Add Python to PATH"
echo   3. Restart your computer
echo   4. Run this script again
echo.
echo Press any key to exit...
pause >nul
exit /b 1

:python_found
echo [OK] Using: %PYTHON_CMD%
echo.

:: ===== CHECK/CREATE VIRTUAL ENVIRONMENT =====
if not exist ".venv\Scripts\python.exe" (
    echo [*] Creating virtual environment...
    %PYTHON_CMD% -m venv .venv
    if !errorlevel! neq 0 (
        echo [X] Failed to create virtual environment.
        echo     Try running: %PYTHON_CMD% -m pip install --user virtualenv
        echo.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created
) else (
    echo [OK] Virtual environment exists
)

:: ===== ACTIVATE VIRTUAL ENVIRONMENT =====
echo [*] Activating virtual environment...
if not exist ".venv\Scripts\activate.bat" (
    echo [X] ERROR: activate.bat not found!
    echo     Delete the .venv folder and run this script again.
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat

:: Verify activation worked
where python 2>nul | findstr /i ".venv" >nul
if %errorlevel% neq 0 (
    echo [!] Warning: Virtual environment may not be properly activated
)
echo [OK] Environment activated
echo.

:: ===== CHECK/INSTALL DEPENDENCIES =====
echo [*] Checking dependencies...

:: Quick check if PyQt6 is installed
python -c "import PyQt6; print('    PyQt6 version:', PyQt6.QtCore.PYQT_VERSION_STR)" 2>nul
if %errorlevel% neq 0 (
    echo.
    echo [*] Installing dependencies (this may take 5-10 minutes)...
    echo     Please wait...
    echo.
    
    :: Upgrade pip first
    python -m pip install --upgrade pip --quiet
    
    :: Install requirements
    python -m pip install -r requirements.txt
    
    if !errorlevel! neq 0 (
        echo.
        echo [X] ERROR: Failed to install dependencies!
        echo.
        echo Possible solutions:
        echo   1. Check your internet connection
        echo   2. Try running as Administrator
        echo   3. Delete .venv folder and try again
        echo.
        pause
        exit /b 1
    )
    echo.
    echo [OK] Dependencies installed successfully
) else (
    echo [OK] Dependencies are installed
)

echo.
echo ===========================================
echo     Starting Application...
echo ===========================================
echo.
echo (If the window closes immediately, check run_error.log)
echo.

:: ===== RUN THE APPLICATION =====
:: Run Python with error output to both console and file
python pgta_report_generator.py 2>&1 | tee run_error.log

:: If tee doesn't exist, use simple redirect
if %errorlevel% neq 0 (
    python pgta_report_generator.py
)

:: Check if there was an error
if %errorlevel% neq 0 (
    echo.
    echo ===========================================
    echo [X] APPLICATION ERROR
    echo ===========================================
    echo.
    if exist run_error.log (
        echo Error log contents:
        echo -------------------------------------------
        type run_error.log
        echo -------------------------------------------
    )
    echo.
    echo Troubleshooting:
    echo   1. Delete .venv folder and run launch.bat again
    echo   2. Check if Python is 64-bit (recommended)
    echo   3. Try: pip install -r requirements.txt --upgrade
    echo.
)

:: ===== CLEANUP =====
call .venv\Scripts\deactivate.bat 2>nul

echo.
echo Application closed.
echo Press any key to exit...
pause >nul
ENDLOCAL

