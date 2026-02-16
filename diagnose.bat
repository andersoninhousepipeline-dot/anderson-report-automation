@echo off
TITLE PGT-A Report Generator - Windows Diagnostic

cd /d "%~dp0"

echo =========================================
echo   PGT-A Report Generator - DIAGNOSTIC
echo =========================================
echo.

echo [*] System Information...
echo   OS: %OS%
echo   Computer: %COMPUTERNAME%
echo   User: %USERNAME%
echo   Date: %DATE%
echo   Time: %TIME%
echo.

echo [*] Python Detection...
python --version >nul 2>&1
if errorlevel 1 (
    echo   Python: NOT FOUND
    py --version >nul 2>&1
    if errorlevel 1 (
        echo   Py command: NOT FOUND
    ) else (
        echo   Py command: FOUND
        py --version
    )
) else (
    echo   Python: FOUND
    python --version
)
echo.

echo [*] Virtual Environment Check...
if exist ".venv" (
    echo   .venv: EXISTS
    if exist ".venv\Scripts\activate.bat" (
        echo   Activation script: EXISTS
        call .venv\Scripts\activate.bat >nul 2>&1
        if errorlevel 1 (
            echo   Activation: FAILED
        ) else (
            echo   Activation: SUCCESS
        )
    ) else (
        echo   Activation script: MISSING
    )
) else (
    echo   .venv: NOT FOUND
)
echo.

echo [*] File Structure Check...
if exist "pgta_report_generator.py" (
    echo   Main app: EXISTS
) else (
    echo   Main app: MISSING
)

if exist "requirements.txt" (
    echo   Requirements: EXISTS
) else (
    echo   Requirements: MISSING
)

if exist "assets\pgta" (
    echo   Assets folder: EXISTS
    dir "assets\pgta" /b >nul 2>&1
    if errorlevel 1 (
        echo   Assets contents: EMPTY
    ) else (
        echo   Assets contents: HAS FILES
    )
) else (
    echo   Assets folder: MISSING
)
echo.

echo [*] Dependency Check (if venv exists)...
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat >nul 2>&1
    if not errorlevel 1 (
        echo   Testing imports...
        
        python -c "import PyQt6; print('   PyQt6: OK')" 2>nul || echo   PyQt6: MISSING
        python -c "import reportlab; print('   ReportLab: OK')" 2>nul || echo   ReportLab: MISSING
        python -c "import pandas; print('   Pandas: OK')" 2>nul || echo   Pandas: MISSING
        python -c "import docx; print('   python-docx: OK')" 2>nul || echo   python-docx: MISSING
        python -c "import easyocr; print('   EasyOCR: OK')" 2>nul || echo   EasyOCR: MISSING
        python -c "import pdfplumber; print('   pdfplumber: OK')" 2>nul || echo   pdfplumber: MISSING
    )
)
echo.

echo [*] Path Information...
echo   Current directory: %CD%
echo   Script location: %~dp0
echo   Python path: 
where python 2>nul || echo     Not found
echo   Py path:
where py 2>nul || echo     Not found
echo.

echo [*] Common Issues Check...
REM Check for common Windows issues
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\python.exe" >nul 2>&1
if errorlevel 1 (
    echo   Python registry entry: MISSING
) else (
    echo   Python registry entry: FOUND
)

REM Check for Visual C++ Redistributable
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90250000-6000-11D3-8CFE-0150048383C9}" >nul 2>&1
if errorlevel 1 (
    reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{1D8E6291-B0D5-35EC-8441-6616F567A0F7}" >nul 2>&1
    if errorlevel 1 (
        echo   Visual C++ 2017+: MAYBE MISSING
    ) else (
        echo   Visual C++ 2017+: FOUND
    )
) else (
    echo   Visual C++ 2017+: FOUND
)
echo.

echo =========================================
echo   DIAGNOSTIC COMPLETE
echo =========================================
echo.
echo Save this output and share if you need support.
echo.
pause
