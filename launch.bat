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

REM Create virtual environment if needed
if not exist ".venv\Scripts\python.exe" (
    echo [*] Creating virtual environment...
    %PYCMD% -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment
        pause
        exit /b 1
    )
    echo [OK] Created
    echo.
)

REM Activate virtual environment
echo [*] Activating environment...
call .venv\Scripts\activate.bat
echo [OK] Activated
echo.

REM Check if dependencies need to be installed
python -c "import PyQt6" >nul 2>&1
if errorlevel 1 (
    echo [*] Installing dependencies...
    echo     This may take several minutes on first run.
    echo.
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo [ERROR] Failed to install dependencies
        echo Try deleting .venv folder and running again
        pause
        exit /b 1
    )
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

if errorlevel 1 (
    echo.
    echo [ERROR] Application crashed!
    echo.
    if exist run_error.log (
        type run_error.log
    )
    echo.
)

echo.
echo Application closed.
pause


