@echo off
TITLE PGT-A Report Generator
echo ===================================
echo PGT-A Report Generator (Windows)
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

echo Using Python command: %PYTHON_CMD%

:: Check dependencies
echo Checking dependencies...
%PYTHON_CMD% -c "import reportlab, docx, PyQt6, pandas, openpyxl, PIL, pdfplumber, PyPDF2" >nul 2>&1
if %errorlevel% neq 0 (
    echo Missing dependencies detected!
    echo Upgrading pip...
    %PYTHON_CMD% -m pip install --upgrade pip
    
    echo Installing required packages...
    %PYTHON_CMD% -m pip install -r requirements_app.txt
    
    if %errorlevel% neq 0 (
        echo.
        echo Error: Failed to install dependencies.
        echo This often happens when PyQt6 cannot be installed.
        echo.
        echo TRY THIS MANUALLY: 
        echo Open Command Prompt and type: %PYTHON_CMD% -m pip install PyQt6 --user
        echo.
        pause
        exit /b
    )
)

:: Check for wkhtmltopdf
where wkhtmltopdf >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [WARNING] wkhtmltopdf not found in PATH!
    echo PDF generation will not work without it.
    echo Please download and install it from: https://wkhtmltopdf.org/downloads.html
    echo After installing, make sure to add the 'bin' folder to your System PATH.
    echo.
)

echo.
echo Starting PGT-A Report Generator...
echo.
%PYTHON_CMD% pgta_report_generator.py

echo.
echo Application closed.
pause
