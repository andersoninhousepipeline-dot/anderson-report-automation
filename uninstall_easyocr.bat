@echo off
TITLE Uninstall EasyOCR and PyTorch
cd /d "%~dp0"

echo.
echo =========================================
echo   Uninstall EasyOCR and PyTorch
echo =========================================
echo.
echo This will remove EasyOCR and PyTorch from your installation.
echo TRF verification has been removed from the application.
echo.
pause

if not exist ".venv" (
    echo [ERROR] No virtual environment found
    echo Please run setup.bat first
    pause
    exit /b 1
)

echo [*] Activating virtual environment...
call .venv\Scripts\activate.bat

echo.
echo [*] Uninstalling EasyOCR and PyTorch...
python -m pip uninstall easyocr torch torchvision torchaudio -y

if errorlevel 1 (
    echo [WARN] Some packages may not have been installed
) else (
    echo [OK] EasyOCR and PyTorch removed
)

echo.
echo =========================================
echo   Cleanup Complete!
echo =========================================
echo.
echo EasyOCR and PyTorch have been removed.
echo Your application will now run without these dependencies.
echo.
pause
