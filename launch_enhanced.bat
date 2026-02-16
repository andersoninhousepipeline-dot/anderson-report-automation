@echo off
TITLE PGT-A Report Generator - Enhanced Windows Launcher

cd /d "%~dp0"

echo =========================================
echo   PGT-A Report Generator - Enhanced
echo =========================================
echo.

REM Set environment variables for better Windows compatibility
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

REM Check for Python with better error handling
echo [*] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    py --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python is not installed or not in PATH!
        echo.
        echo Please:
        echo   1. Go to https://www.python.org/downloads/
        echo   2. Download and install Python 3.10+
        echo   3. CHECK "Add Python to PATH" during installation
        echo   4. Restart Command Prompt and try again
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

REM Check Python architecture (must be 64-bit for PyQt6)
echo [*] Checking Python architecture...
%PYCMD% -c "import struct; bits=struct.calcsize('P')*8; print(f'  Architecture: {bits}-bit'); exit(0 if bits==64 else 1)"
if errorlevel 1 (
    echo [ERROR] 32-bit Python detected! PyQt6 requires 64-bit Python.
    echo         Please uninstall and reinstall the 64-bit version.
    pause
    exit /b 1
)
echo.

REM Check if virtual environment exists
if not exist ".venv" (
    echo [*] Virtual environment not found. Running setup first...
    call setup.bat
    if errorlevel 1 (
        echo [ERROR] Setup failed. Please check the error messages above.
        pause
        exit /b 1
    )
)

REM Activate virtual environment
echo [*] Activating virtual environment...
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
    echo [OK] Virtual environment activated
) else (
    echo [ERROR] Virtual environment activation script not found!
    echo Please run setup.bat first.
    pause
    exit /b 1
)
echo.

REM Check if core dependencies are installed
echo [*] Checking dependencies...
.venv\Scripts\python.exe -c "import PyQt6" >nul 2>&1
if errorlevel 1 (
    echo [WARN] PyQt6 not found. Installing dependencies...
    .venv\Scripts\python.exe -m pip install --upgrade pip
    .venv\Scripts\python.exe -m pip install PyQt6>=6.6.0 PyQt6-Qt6>=6.6.0 reportlab>=4.0.0 PyPDF2>=3.0.0 pdfplumber>=0.10.0 python-docx>=1.0.0 pandas>=2.0.0 openpyxl>=3.1.0 Pillow>=10.0.0 numpy>=1.24.0 requests>=2.31.0
    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies
        pause
        exit /b 1
    )
)

REM Check if main application file exists
if not exist "pgta_report_generator.py" (
    echo [ERROR] Main application file not found!
    echo Make sure you're running this from the correct directory.
    pause
    exit /b 1
)

REM Check if assets directory exists
if not exist "assets\pgta" (
    echo [ERROR] Assets directory not found!
    echo Make sure the assets/pgta folder exists with required files.
    pause
    exit /b 1
)

echo.
echo =========================================
echo Starting PGT-A Report Generator...
echo =========================================
echo.

REM Start the application with error handling
.venv\Scripts\python.exe pgta_report_generator.py 2> run_error.log

REM Check exit code
if errorlevel 1 (
    echo.
    echo [ERROR] Application exited with error code %errorlevel%
    echo.
    echo Common issues:
    echo   - Missing Visual C++ Redistributable (for EasyOCR/PyTorch)
    echo   - PyQt6 installation issues
    echo   - Missing assets files
    echo.
    echo Try:
    echo   1. Install Visual C++ Redistributable: https://aka.ms/vs/17/release/vc_redist.x64.exe
    echo   2. Reinstall dependencies: pip uninstall PyQt6 PyQt6-Qt6 -y && pip install PyQt6
    echo   3. Check assets folder exists
    echo.
) else (
    echo.
    echo =========================================
    echo Application closed successfully.
    echo =========================================
)

pause
