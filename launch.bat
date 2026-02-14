@echo off
TITLE PGT-A Report Generator

cd /d "%~dp0"

echo =========================================
echo   PGT-A Report Generator
echo =========================================
echo.

REM Set environment variables for Windows compatibility
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

REM --- Step 1: Find Python ---
echo [*] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    py --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python is not installed or not in PATH!
        echo.
        echo Please:
        echo   1. Go to https://www.python.org/downloads/
        echo   2. Download and install Python 3.10+ (64-bit)
        echo   3. CHECK "Add Python to PATH" during installation
        echo   4. Restart your computer and try again
        echo.
        pause
        exit /b 1
    ) else (
        set PYCMD=py
    )
) else (
    set PYCMD=python
)

echo [OK] Python found:
%PYCMD% --version
echo.

REM --- Step 2: Check Python Architecture (must be 64-bit) ---
echo [*] Checking Python architecture...
%PYCMD% -c "import struct; bits=struct.calcsize('P')*8; print(f'  Python Architecture: {bits}-bit'); exit(0 if bits==64 else 1)"
if errorlevel 1 (
    echo.
    echo [ERROR] 32-bit Python detected! PyQt6 requires 64-bit Python.
    echo         Please uninstall Python and reinstall the 64-bit version.
    echo         Download: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)
echo.

REM --- Step 3: Create virtual environment if missing ---
if not exist ".venv" (
    echo [*] Creating virtual environment...
    %PYCMD% -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        echo         Try running as Administrator or check antivirus settings.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created.
)
echo.

REM --- Step 4: Activate virtual environment ---
echo [*] Activating virtual environment...
if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] Virtual environment is broken. Deleting and recreating...
    rmdir /s /q .venv 2>nul
    timeout /t 2 >nul
    %PYCMD% -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to recreate virtual environment.
        pause
        exit /b 1
    )
)
call .venv\Scripts\activate.bat
echo [OK] Virtual environment activated.
echo.

REM --- Step 5: Check & install dependencies ---
.venv\Scripts\python.exe -c "import PyQt6" >nul 2>&1
if errorlevel 1 (
    echo [*] Installing core packages (first run)...
    .venv\Scripts\python.exe -m pip install --upgrade pip >nul 2>&1
    .venv\Scripts\python.exe -m pip install PyQt6>=6.6.0 PyQt6-Qt6>=6.6.0 reportlab>=4.0.0 PyPDF2>=3.0.0 pdfplumber>=0.10.0 python-docx>=1.0.0 pandas>=2.0.0 openpyxl>=3.1.0 Pillow>=10.0.0 numpy>=1.24.0 requests>=2.31.0
    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies.
        echo         Check your internet connection and try again.
        echo         If PyQt6 fails, install Visual C++ Redistributable:
        echo         https://aka.ms/vs/17/release/vc_redist.x64.exe
        pause
        exit /b 1
    )
    echo.
    echo [OK] Core packages installed. Starting app...
    echo     Note: EasyOCR (optional) will be installed on next setup run.
)
echo.

REM --- Step 6: Verify main file exists ---
if not exist "pgta_report_generator.py" (
    echo [ERROR] pgta_report_generator.py not found!
    echo         Make sure you are running this from the correct directory.
    pause
    exit /b 1
)

REM --- Step 7: Launch Application ---
echo =========================================
echo   Starting application...
echo =========================================
echo.

.venv\Scripts\python.exe pgta_report_generator.py 2> run_error.log

REM --- Step 8: Check exit status ---
if errorlevel 1 (
    echo.
    echo =========================================
    echo   [ERROR] Application crashed!
    echo =========================================
    echo.
    echo Error details:
    type run_error.log
    echo.
    echo.
    echo Common fixes:
    echo   1. Install Visual C++ Redistributable:
    echo      https://aka.ms/vs/17/release/vc_redist.x64.exe
    echo   2. Delete .venv folder and run this script again
    echo   3. Ensure 64-bit Python is installed
    echo.
) else (
    echo.
    echo =========================================
    echo   Application closed successfully.
    echo =========================================
)

pause
