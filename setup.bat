@echo off
:: ===========================================
:: PGT-A Report Generator - Quick Setup Script
:: ===========================================
:: Run this once to set up the environment

TITLE PGT-A Report Generator - Setup
chcp 65001 >nul 2>&1

echo ===========================================
echo   PGT-A Report Generator - Setup
echo ===========================================
echo.

:: Change to script directory
cd /d "%~dp0"

:: Check for Python
set PYTHON_CMD=
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python
) else (
    py --version >nul 2>&1
    if %errorlevel% equ 0 (
        set PYTHON_CMD=py
    )
)

if "%PYTHON_CMD%"=="" (
    echo [X] Python is not installed!
    echo.
    echo Please install Python 3.10 or higher:
    echo   1. Download from https://www.python.org/downloads/
    echo   2. CHECK "Add Python to PATH" during installation
    echo   3. Restart computer and run this script again
    echo.
    pause
    exit /b 1
)

echo [OK] Python found
%PYTHON_CMD% --version
echo.

:: Remove old venv if exists
if exist ".venv" (
    echo [*] Removing old virtual environment...
    rmdir /s /q .venv
)

:: Create new venv
echo [*] Creating virtual environment...
%PYTHON_CMD% -m venv .venv
if %errorlevel% neq 0 (
    echo [X] Failed to create virtual environment
    pause
    exit /b 1
)

:: Activate
call .venv\Scripts\activate.bat

:: Upgrade pip
echo [*] Upgrading pip...
python -m pip install --upgrade pip -q

:: Install dependencies
echo.
echo [*] Installing dependencies...
echo     (This may take 5-10 minutes on first install)
echo.
python -m pip install -r requirements.txt

if %errorlevel% equ 0 (
    echo.
    echo ===========================================
    echo [OK] Setup Complete!
    echo ===========================================
    echo.
    echo To run the application:
    echo   Double-click launch.bat
    echo.
    echo Or from command line:
    echo   .venv\Scripts\activate
    echo   python pgta_report_generator.py
    echo.
) else (
    echo.
    echo [X] Setup failed. Check errors above.
    pause
    exit /b 1
)

call deactivate 2>nul
pause
