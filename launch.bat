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
    echo Installing required packages...
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo.
        echo [ERROR] Package installation failed.
        echo Please check your internet connection and try again.
        pause
        exit
    )
    echo.
    echo Packages installed. Verifying PyQt6...
    python -c "import PyQt6" >nul 2>&1
    if %errorlevel% neq 0 (
        echo.
        echo [ERROR] PyQt6 could not be imported after install.
        echo Your system may need the Microsoft Visual C++ Redistributable.
        echo Download from: https://aka.ms/vs/17/release/vc_redist.x64.exe
        echo.
        pause
        exit
    )
    echo PyQt6 ready. Starting app...
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
