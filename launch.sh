#!/bin/bash
# PGT-A Report Generator Launcher Script

echo "==================================="
echo "PGT-A Report Generator (Linux)"
echo "==================================="
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is not installed"
    echo "Please install Python 3.8 or higher"
    exit 1
fi

# Create Virtual Environment if it doesn't exist
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv .venv
    if [ $? -ne 0 ]; then
        echo "Error: Failed to create virtual environment"
        exit 1
    fi
fi

# Activate Virtual Environment
echo "Activating environment..."
source .venv/bin/activate

# Install/Update dependencies
echo "Checking/Installing dependencies..."
pip install --upgrade pip
pip install -r requirements.txt

if [ $? -ne 0 ]; then
    echo ""
    echo "Error: Failed to install dependencies"
    exit 1
fi

echo ""
echo "Starting PGT-A Report Generator..."
echo ""

# Launch the application
python3 pgta_report_generator.py 2>run_error.log

if [ $? -ne 0 ]; then
    echo ""
    echo "[ERROR] The application failed to start."
    echo "Detailed error:"
    cat run_error.log
    echo ""
fi

deactivate
echo ""
echo "Application closed."
