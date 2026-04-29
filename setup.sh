#!/bin/zsh

echo "============================================"
echo " CA3 Toolkit - Environment Setup"
echo "============================================"
echo

# Check Python is available
if ! command -v python3 &>/dev/null; then
    echo "ERROR: Python 3 is not installed or not in PATH."
    echo "       Install via Homebrew:  brew install python"
    echo "       Or download from:      https://www.python.org/downloads/"
    exit 1
fi

echo "Found: $(python3 --version)"
echo

# Create virtual environment
if [ -d ".venv" ]; then
    echo "Virtual environment already exists, skipping creation."
else
    echo "Creating virtual environment..."
    python3 -m venv .venv
    if [ $? -ne 0 ]; then
        echo "ERROR: Failed to create virtual environment."
        exit 1
    fi
    echo "Done."
fi
echo

# Activate virtual environment
echo "Activating virtual environment..."
source .venv/bin/activate
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to activate virtual environment."
    exit 1
fi
echo "Done."
echo

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip --quiet
echo "Done."
echo

# Install dependencies
echo "Installing dependencies from requirements.txt..."
pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to install dependencies."
    exit 1
fi
echo

echo "============================================"
echo " Setup complete!"
echo "============================================"
echo
echo "To activate the environment in future sessions, run:"
echo "    source .venv/bin/activate"
echo
echo "Workflow:"
echo "  1. python extract_to_excel.py"
echo "  2. Open marks.xlsx and fill in marks"
echo "  3. python fill_marks_and_export.py"
echo

exec $SHELL
