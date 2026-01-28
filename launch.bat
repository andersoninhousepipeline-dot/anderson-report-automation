@echo off
TITLE ADVAT Report Generator
echo ===================================
echo ADVAT Report Generator (Windows)
echo ===================================
echo.

:: Check for Python
set PYTHON_CMD=
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python
) else (
    python3 --version >nul 2>&1
    if %errorlevel% equ 0 (
        set PYTHON_CMD=python3
    ) else (
        py --version >nul 2>&1
        if %errorlevel% equ 0 (
            set PYTHON_CMD=py
        )
    )
)

if "%PYTHON_CMD%"=="" (
    echo Error: Python was not found on your system.
    echo.
    echo 1. Please install Python 3.10+ from https://www.python.org/
    echo 2. IMPORTANT: During installation, CHECK the box "Add Python to PATH".
    echo 3. Restart your computer after installation.
    echo.
    pause
    exit /b
)

:: Create Virtual Environment if it doesn't exist
if not exist ".venv" (
    echo Creating virtual environment...
    %PYTHON_CMD% -m venv .venv
    if %errorlevel% neq 0 (
        echo Error: Failed to create virtual environment.
        pause
        exit /b
    )
)

:: Activate Virtual Environment
echo Activating environment...
call .venv\Scripts\activate.bat

:: Install/Update dependencies
echo Checking/Installing dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo Error: Failed to install dependencies.
    pause
    exit /b
)

echo.
echo Starting ADVAT Report Generator...
echo.
python pgta_report_generator.py 2>run_error.log

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] The application failed to start.
    echo Detailed error:
    type run_error.log
    echo.
    pause
)

deactivate
echo.
echo Application closed.
pause
