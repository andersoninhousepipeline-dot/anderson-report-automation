@echo off
TITLE PGT-A Report Generator - Windows Launcher
cd /d "%~dp0"

REM ============================================
REM  PGT-A Report Generator - Windows Launcher
REM  Enhanced with comprehensive error handling
REM ============================================

echo.
echo =========================================
echo   PGT-A Report Generator
echo =========================================
echo.

REM Create log file
set LOGFILE=launcher.log
echo [%DATE% %TIME%] Launcher started > %LOGFILE%

REM ============================================
REM  STEP 1: Python Detection
REM ============================================
echo [*] Checking for Python...
echo [%DATE% %TIME%] Checking for Python >> %LOGFILE%

set PYCMD=
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYCMD=python
    echo [OK] Python found: >> %LOGFILE%
    python --version >> %LOGFILE% 2>&1
    python --version
) else (
    py --version >nul 2>&1
    if %errorlevel% equ 0 (
        set PYCMD=py
        echo [OK] Py launcher found: >> %LOGFILE%
        py --version >> %LOGFILE% 2>&1
        py --version
    ) else (
        py -3 --version >nul 2>&1
        if %errorlevel% equ 0 (
            set PYCMD=py -3
            echo [OK] Py -3 launcher found: >> %LOGFILE%
            py -3 --version >> %LOGFILE% 2>&1
            py -3 --version
        )
    )
)

if "%PYCMD%"=="" (
    echo [ERROR] Python not found! >> %LOGFILE%
    echo.
    echo [ERROR] Python is not installed or not in PATH!
    echo.
    echo Please install Python:
    echo   1. Go to: https://www.python.org/downloads/
    echo   2. Download Python 3.10 or newer (64-bit)
    echo   3. During installation, CHECK "Add Python to PATH"
    echo   4. Restart your computer
    echo   5. Run this script again
    echo.
    pause
    exit /b 1
)

echo.

REM ============================================
REM  STEP 2: Architecture Check (64-bit required)
REM ============================================
echo [*] Checking Python architecture...
echo [%DATE% %TIME%] Checking architecture >> %LOGFILE%

%PYCMD% -c "import struct; import sys; bits=struct.calcsize('P')*8; print(f'  Architecture: {bits}-bit'); sys.exit(0 if bits==64 else 1)" 2>>%LOGFILE%
if %errorlevel% neq 0 (
    echo [ERROR] 32-bit Python detected! >> %LOGFILE%
    echo.
    echo [ERROR] 32-bit Python detected!
    echo.
    echo PyQt6 requires 64-bit Python.
    echo.
    echo Please:
    echo   1. Uninstall current Python
    echo   2. Download 64-bit Python from python.org
    echo   3. Install with "Add to PATH" checked
    echo   4. Restart computer and try again
    echo.
    pause
    exit /b 1
)
echo.

REM ============================================
REM  STEP 3: Virtual Environment Setup
REM ============================================
if not exist ".venv" (
    echo [*] Creating virtual environment...
    echo [%DATE% %TIME%] Creating venv >> %LOGFILE%
    %PYCMD% -m venv .venv 2>>%LOGFILE%
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment >> %LOGFILE%
        echo [ERROR] Failed to create virtual environment
        echo Check %LOGFILE% for details
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created
)

echo [*] Activating virtual environment...
echo [%DATE% %TIME%] Activating venv >> %LOGFILE%
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat 2>>%LOGFILE%
    echo [OK] Activated
) else (
    echo [ERROR] Activation script not found >> %LOGFILE%
    echo [ERROR] Virtual environment is corrupted
    echo.
    echo Please delete the .venv folder and run this script again
    pause
    exit /b 1
)
echo.

REM ============================================
REM  STEP 4: Dependency Check
REM ============================================
echo [*] Checking dependencies...
echo [%DATE% %TIME%] Checking dependencies >> %LOGFILE%

.venv\Scripts\python.exe -c "import PyQt6" >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARN] PyQt6 not found, installing dependencies...
    echo [%DATE% %TIME%] Installing dependencies >> %LOGFILE%
    
    .venv\Scripts\python.exe -m pip install --upgrade pip >>%LOGFILE% 2>&1
    .venv\Scripts\python.exe -m pip install PyQt6>=6.6.0 PyQt6-Qt6>=6.6.0 reportlab>=4.0.0 PyPDF2>=3.0.0 pdfplumber>=0.10.0 python-docx>=1.0.0 pandas>=2.0.0 openpyxl>=3.1.0 Pillow>=10.0.0 numpy>=1.24.0 requests>=2.31.0 >>%LOGFILE% 2>&1
    
    if %errorlevel% neq 0 (
        echo [ERROR] Dependency installation failed >> %LOGFILE%
        echo.
        echo [ERROR] Failed to install dependencies
        echo.
        echo This may be caused by:
        echo   1. No internet connection
        echo   2. Firewall blocking pip
        echo   3. Insufficient disk space
        echo.
        echo Check %LOGFILE% for details
        pause
        exit /b 1
    )
    echo [OK] Dependencies installed
) else (
    echo [OK] Dependencies found
)
echo.

REM ============================================
REM  STEP 5: Asset Check
REM ============================================
echo [*] Checking assets...
echo [%DATE% %TIME%] Checking assets >> %LOGFILE%

if not exist "assets\pgta" (
    echo [ERROR] Assets folder missing >> %LOGFILE%
    echo [ERROR] Assets folder not found!
    echo.
    echo The application requires the assets/pgta folder.
    echo Please ensure it exists in the same directory as this script.
    pause
    exit /b 1
)
echo [OK] Assets found
echo.

REM ============================================
REM  STEP 6: Launch Application
REM ============================================
echo =========================================
echo   Starting Application...
echo =========================================
echo.
echo [%DATE% %TIME%] Launching application >> %LOGFILE%

.venv\Scripts\python.exe pgta_report_generator.py 2>run_error.log

REM ============================================
REM  STEP 7: Check Exit Status
REM ============================================
if %errorlevel% neq 0 (
    echo. >> %LOGFILE%
    echo [ERROR] Application crashed with exit code %errorlevel% >> %LOGFILE%
    echo.
    echo =========================================
    echo   [ERROR] Application crashed!
    echo =========================================
    echo.
    echo Exit code: %errorlevel%
    echo.
    echo Error logs created:
    echo   - run_error.log
    echo   - startup.log (if created)
    echo   - launcher.log
    echo.
    echo Common solutions:
    echo   1. Check startup.log for detailed error info
    echo   2. Run diagnose.bat for system diagnostics
    echo   3. Delete .venv folder and run this script again
    echo.
    if exist "run_error.log" (
        echo Last error from run_error.log:
        type run_error.log
    )
    echo.
) else (
    echo [%DATE% %TIME%] Application closed normally >> %LOGFILE%
    echo.
    echo =========================================
    echo   Application closed successfully
    echo =========================================
)

echo.
pause
