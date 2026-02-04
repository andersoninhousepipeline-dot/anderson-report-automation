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

REM Install dependencies
echo [*] Installing dependencies...
echo     This may take 5-15 minutes. Please wait...
echo.
python -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install dependencies
    echo.
    echo Try:
    echo   1. Check internet connection
    echo   2. Run as Administrator
    echo   3. Delete .venv folder and try again
    echo.
    pause
    exit /b 1
)

echo.
echo ===========================================
echo   SETUP COMPLETE!
echo ===========================================
echo.
echo To run the application, double-click: launch.bat
echo.
pause


