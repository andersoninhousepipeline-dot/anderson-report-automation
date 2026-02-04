#!/bin/bash
# ===========================================
# PGT-A Report Generator - Linux/Mac Launcher
# ===========================================
# This script sets up and runs the PGT-A Report Generator

set -e  # Exit on error

echo "==========================================="
echo "  PGT-A Report Generator (Linux/Mac)"
echo "==========================================="
echo ""

# Get script directory
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Error: Python 3 is not installed"
    echo ""
    echo "Please install Python 3.8 or higher:"
    echo "  Ubuntu/Debian: sudo apt install python3 python3-pip python3-venv"
    echo "  Fedora:        sudo dnf install python3 python3-pip"
    echo "  macOS:         brew install python3"
    echo ""
    exit 1
fi

# Check Python version
PYTHON_VERSION=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
echo "✓ Python $PYTHON_VERSION found"

# Create Virtual Environment if it doesn't exist
if [ ! -d ".venv" ]; then
    echo ""
    echo "📦 Creating virtual environment..."
    python3 -m venv .venv
    if [ $? -ne 0 ]; then
        echo "❌ Error: Failed to create virtual environment"
        echo "Try: sudo apt install python3-venv"
        exit 1
    fi
    echo "✓ Virtual environment created"
fi

# Activate Virtual Environment
echo ""
echo "🔄 Activating environment..."
source .venv/bin/activate

# Upgrade pip
pip install --upgrade pip -q

# Check if dependencies are installed
if ! python3 -c "import PyQt6" 2>/dev/null; then
    echo ""
    echo "📥 Installing dependencies (first run may take a few minutes)..."
    echo "   This includes PyQt6, EasyOCR, and other packages..."
    echo ""
    pip install -r requirements.txt
    
    if [ $? -ne 0 ]; then
        echo ""
        echo "❌ Error: Failed to install dependencies"
        echo ""
        echo "Try installing manually:"
        echo "  source .venv/bin/activate"
        echo "  pip install -r requirements.txt"
        exit 1
    fi
    echo ""
    echo "✓ All dependencies installed"
else
    echo "✓ Dependencies already installed"
fi

echo ""
echo "🚀 Starting PGT-A Report Generator..."
echo "==========================================="
echo ""

# Launch the application
python3 pgta_report_generator.py 2>run_error.log

EXIT_CODE=$?
if [ $EXIT_CODE -ne 0 ]; then
    echo ""
    echo "❌ [ERROR] The application crashed (exit code: $EXIT_CODE)"
    echo ""
    echo "Error details:"
    echo "-------------------------------------------"
    cat run_error.log
    echo "-------------------------------------------"
    echo ""
    echo "Common fixes:"
    echo "  1. Update packages: pip install -r requirements.txt --upgrade"
    echo "  2. Reinstall venv: rm -rf .venv && ./launch.sh"
    echo ""
fi

deactivate 2>/dev/null || true
echo ""
echo "Application closed."
