@echo off
TITLE PGT-A Report Generator
chcp 65001 >nul 2>&1
echo ===========================================
echo   PGT-A Report Generator (Windows)
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
    echo [X] Error: Python was not found on your system.
    echo.
    echo Please install Python 3.10 or higher:
    echo   1. Download from https://www.python.org/downloads/
    echo   2. IMPORTANT: Check "Add Python to PATH" during installation
    echo   3. Restart your computer after installation
    echo.
    pause
    exit /b 1
)

:: Check Python version and architecture
echo.
%PYTHON_CMD% -c "import sys; v=sys.version_info; print(f'[OK] Python {v.major}.{v.minor}.{v.micro} found')"
%PYTHON_CMD% -c "import struct; arch=struct.calcsize('P')*8; print(f'[OK] {arch}-bit architecture')"

:: Check for 64-bit (recommended)
%PYTHON_CMD% -c "import struct; exit(0 if struct.calcsize('P')*8==64 else 1)" >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [!] WARNING: 32-bit Python detected.
    echo     64-bit Python is recommended for best performance.
    echo     Download from: https://www.python.org/downloads/
    echo.
    set /p "continue_32bit=Continue anyway? (y/n): "
    if /i not "%continue_32bit%"=="y" exit /b 1
)
echo.

:: Create Virtual Environment if it doesn't exist
if not exist ".venv" (
    echo [*] Creating virtual environment...
    %PYTHON_CMD% -m venv .venv
    if %errorlevel% neq 0 (
        echo [X] Error: Failed to create virtual environment.
        echo     Try: %PYTHON_CMD% -m pip install virtualenv
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created
)

:: Activate Virtual Environment
echo [*] Activating environment...
call .venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo [X] Error: Failed to activate virtual environment.
    pause
    exit /b 1
)

:: Upgrade pip quietly
python -m pip install --upgrade pip -q

:: Check if PyQt6 is installed (indicator that deps are installed)
python -c "import PyQt6" >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [*] Installing dependencies (first run may take several minutes)...
    echo     This includes PyQt6, EasyOCR, torch, and other packages...
    echo.
    python -m pip install -r requirements.txt
    
    if %errorlevel% neq 0 (
        echo.
        echo [X] Error: Failed to install dependencies.
        echo.
        echo Try manually:
        echo   1. Open Command Prompt
        echo   2. cd "%~dp0"
        echo   3. .venv\Scripts\activate
        echo   4. pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
    echo.
    echo [OK] All dependencies installed
) else (
    echo [OK] Dependencies already installed
)

echo.
echo ===========================================
echo [*] Starting PGT-A Report Generator...
echo ===========================================
echo.

:: Launch the application
python pgta_report_generator.py 2> run_error.log

if %errorlevel% neq 0 (
    echo.
    echo [X] ERROR: The application crashed!
    echo.
    echo Error details:
    echo -------------------------------------------
    type run_error.log
    echo -------------------------------------------
    echo.
    echo Common fixes:
    echo   1. Update packages: pip install -r requirements.txt --upgrade
    echo   2. Reinstall venv: Delete .venv folder and run launch.bat again
    echo.
    pause
)

call deactivate 2>nul
echo.
echo Application closed.
pause
