#!/bin/bash
# ===========================================
# PGT-A Report Generator - Quick Setup Script
# ===========================================
# Run this once to set up the environment

echo "==========================================="
echo "  PGT-A Report Generator - Setup"
echo "==========================================="
echo ""

# Get script directory
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed!"
    echo ""
    echo "Install Python first:"
    echo "  Ubuntu/Debian: sudo apt install python3 python3-pip python3-venv"
    echo "  Fedora:        sudo dnf install python3 python3-pip"
    echo "  macOS:         brew install python3"
    exit 1
fi

echo "✓ Python 3 found: $(python3 --version)"
echo ""

# Remove old venv if exists
if [ -d ".venv" ]; then
    echo "🗑️  Removing old virtual environment..."
    rm -rf .venv
fi

# Create new venv
echo "📦 Creating virtual environment..."
python3 -m venv .venv
source .venv/bin/activate

# Upgrade pip
echo "📥 Upgrading pip..."
pip install --upgrade pip -q

# Install dependencies
echo ""
echo "📥 Installing dependencies..."
echo "   (This may take 5-10 minutes on first install)"
echo ""
pip install -r requirements.txt

if [ $? -eq 0 ]; then
    echo ""
    echo "==========================================="
    echo "✅ Setup Complete!"
    echo "==========================================="
    echo ""
    echo "To run the application:"
    echo "  ./launch.sh"
    echo ""
    echo "Or manually:"
    echo "  source .venv/bin/activate"
    echo "  python3 pgta_report_generator.py"
    echo ""
else
    echo ""
    echo "❌ Setup failed. Check errors above."
    exit 1
fi

deactivate
