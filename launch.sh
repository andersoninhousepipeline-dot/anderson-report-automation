#!/bin/bash
# PGT-A Report Generator Launcher Script

echo "==================================="
echo "PGT-A Report Generator"
echo "==================================="
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is not installed"
    echo "Please install Python 3.8 or higher"
    exit 1
fi

# Check Python version
PYTHON_VERSION=$(python3 -c 'import sys; print(".".join(map(str, sys.version_info[:2])))')
echo "Python version: $PYTHON_VERSION"

# Check if dependencies are installed
echo "Checking dependencies..."
python3 -c "import reportlab, docx, PyQt6, pandas, openpyxl, PIL, pdfplumber" 2>/dev/null

if [ $? -ne 0 ]; then
    echo ""
    echo "Missing dependencies detected!"
    echo "Installing required packages..."
    pip install -r requirements_app.txt
    
    if [ $? -ne 0 ]; then
        echo "Error: Failed to install dependencies"
        exit 1
    fi
fi

echo ""
echo "Starting PGT-A Report Generator..."
echo ""

# Launch the application
python3 pgta_report_generator.py

echo ""
echo "Application closed."
