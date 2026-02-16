@echo off
TITLE PGT-A Report Generator - Enhanced Setup
cd /d "%~dp0"

echo.
echo ===========================================
echo   PGT-A Report Generator - ENHANCED SETUP
echo ===========================================
echo.

REM Set environment variables for better Windows compatibility
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

REM Check for Python with detailed feedback
echo [*] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    py --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python is not installed or not in PATH!
        echo.
        echo SOLUTION:
        echo   1. Go to https://www.python.org/downloads/
        echo   2. Download Python 3.10 or higher (64-bit)
        echo   3. During installation, CHECK "Add Python to PATH"
        echo   4. Restart your computer
        echo   5. Run this setup again
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

REM Check Python version (need 3.8+)
echo [*] Checking Python version...
for /f "tokens=2" %%i in ('%PYCMD% --version 2^>^&1') do set PYTHON_VERSION=%%i
echo   Python version: %PYTHON_VERSION%

REM Check 64-bit architecture (required for PyQt6)
echo [*] Checking Python architecture...
%PYCMD% -c "import struct; bits=struct.calcsize('P')*8; print(f'  Architecture: {bits}-bit'); exit(0 if bits==64 else 1)"
if errorlevel 1 (
    echo [ERROR] 32-bit Python detected! PyQt6 requires 64-bit Python.
    echo         Please uninstall Python and reinstall the 64-bit version.
    echo         Download: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo.

REM Remove old virtual environment with better error handling
if exist ".venv" (
    echo [*] Removing old virtual environment...
    rmdir /s /q .venv 2>nul
    if exist ".venv" (
        echo [WARN] Some files could not be removed. Trying again...
        timeout /t 2 >nul
        rmdir /s /q .venv 2>nul
    )
    if exist ".venv" (
        echo [ERROR] Cannot remove old virtual environment
        echo Please close any running applications and try again
        pause
        exit /b 1
    )
    echo [OK] Old environment removed
    echo.
)

REM Create new virtual environment
echo [*] Creating virtual environment...
%PYCMD% -m venv .venv
if errorlevel 1 (
    echo [ERROR] Failed to create virtual environment
    echo.
    echo Possible causes:
    echo   - Python installation issue
    echo   - Antivirus blocking
    echo   - Permissions issue
    echo.
    echo Try running as Administrator or check antivirus settings.
    pause
    exit /b 1
)
echo [OK] Virtual environment created
echo.

REM Activate virtual environment
echo [*] Activating virtual environment...
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
    echo [OK] Virtual environment activated
) else (
    echo [ERROR] Virtual environment activation failed
    echo Please try running setup again
    pause
    exit /b 1
)
echo.

REM Upgrade pip with better error handling
echo [*] Upgrading pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo [WARN] Pip upgrade failed, continuing with current version...
)
echo.

REM Install core dependencies
echo [*] Installing core dependencies...
echo This may take a few minutes...
echo.

python -m pip install PyQt6>=6.6.0 PyQt6-Qt6>=6.6.0 reportlab>=4.0.0 PyPDF2>=3.0.0 pdfplumber>=0.10.0 python-docx>=1.0.0 pandas>=2.0.0 openpyxl>=3.1.0 Pillow>=10.0.0 numpy>=1.24.0 requests>=2.31.0

if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install core dependencies
    echo.
    echo Troubleshooting:
    echo   1. Check internet connection
    echo   2. Try: pip install --upgrade pip setuptools wheel
    echo   3. Try: pip install PyQt6 --no-cache-dir
    echo   4. Disable antivirus temporarily
    echo.
    pause
    exit /b 1
)
echo [OK] Core dependencies installed
echo.

REM Test core imports
echo [*] Testing core imports...
python -c "
try:
    import PyQt6
    import reportlab
    import pandas
    import docx
    print('[OK] All core imports successful')
except ImportError as e:
    print(f'[ERROR] Import failed: {e}')
    exit(1)
"
if errorlevel 1 (
    echo [ERROR] Core import test failed
    pause
    exit /b 1
)
echo.

REM Install EasyOCR (optional)
echo [*] Installing EasyOCR (optional, for TRF verification)...
echo     This may take several minutes...
echo     If this fails, the app will still work without OCR.
echo.

python -m pip install easyocr>=1.7.0

if errorlevel 1 (
    echo.
    echo [!] EasyOCR installation had issues (this is common on Windows)
    echo.
    echo REASON: EasyOCR requires PyTorch which may need Visual C++ Redistributable
    echo.
    echo SOLUTION: Install Visual C++ Redistributable:
    echo   https://aka.ms/vs/17/release/vc_redist.x64.exe
    echo.
    echo After installing, run setup.bat again to enable EasyOCR.
    echo.
    echo NOTE: The application will work fine without EasyOCR.
) else (
    echo [OK] EasyOCR installed successfully
)
echo.

REM Check assets directory
echo [*] Checking assets directory...
if not exist "assets\pgta" (
    echo [WARN] Assets directory not found!
    echo Please ensure the assets/pgta folder exists with required template files.
    echo.
)

REM Final test
echo [*] Running final test...
python -c "
import sys
import os
try:
    from pgta_template import PGTAReportTemplate
    from pgta_docx_generator import PGTADocxGenerator
    print('[OK] Application modules loaded successfully')
except Exception as e:
    print(f'[ERROR] Module test failed: {e}')
    exit(1)
"
if errorlevel 1 (
    echo [ERROR] Final test failed
    pause
    exit /b 1
)

echo.
echo ===========================================
echo   SETUP COMPLETE!
echo ===========================================
echo.
echo SUCCESS! Your PGT-A Report Generator is ready.
echo.
echo TO RUN: Double-click launch_enhanced.bat
echo.
echo NOTES:
echo   - If you see PyTorch/DLL errors: Install Visual C++ Redistributable
echo   - Link: https://aka.ms/vs/17/release/vc_redist.x64.exe
echo   - Application will work without EasyOCR if needed
echo.
pause
