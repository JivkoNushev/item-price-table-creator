#!/usr/bin/env bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "=== Table Generator ==="

# Check for Python 3
PYTHON=""
for cmd in python3 python; do
    if command -v "$cmd" &>/dev/null; then
        version=$("$cmd" --version 2>&1 | grep -oP '\d+\.\d+')
        major=$(echo "$version" | cut -d. -f1)
        if [ "$major" -ge 3 ]; then
            PYTHON="$cmd"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    echo "Error: Python 3 is required but not found."
    echo "Install it with: sudo apt install python3"
    exit 1
fi

echo "Using $PYTHON ($($PYTHON --version 2>&1))"

# Check for tkinter
if ! "$PYTHON" -c "import tkinter" 2>/dev/null; then
    echo "tkinter is not installed. Attempting to install..."
    if command -v apt &>/dev/null; then
        sudo apt install -y python3-tk
    elif command -v dnf &>/dev/null; then
        sudo dnf install -y python3-tkinter
    elif command -v pacman &>/dev/null; then
        sudo pacman -S --noconfirm tk
    elif command -v brew &>/dev/null; then
        brew install python-tk
    else
        echo "Error: Could not install tkinter automatically."
        echo "Please install the tkinter package for your system."
        exit 1
    fi
fi

# Check for openpyxl and install if missing
if ! "$PYTHON" -c "import openpyxl" 2>/dev/null; then
    echo "openpyxl is not installed. Installing via pip..."
    PIP_INSTALLED=false

    # Try standard install
    if "$PYTHON" -m pip install --quiet openpyxl 2>/dev/null; then
        PIP_INSTALLED=true
    fi

    # Try --user install
    if [ "$PIP_INSTALLED" = false ]; then
        if "$PYTHON" -m pip install --quiet --user openpyxl 2>/dev/null; then
            PIP_INSTALLED=true
        fi
    fi

    # Try --break-system-packages (Debian/Ubuntu externally-managed environments)
    if [ "$PIP_INSTALLED" = false ]; then
        if "$PYTHON" -m pip install --quiet --break-system-packages openpyxl 2>/dev/null; then
            PIP_INSTALLED=true
        fi
    fi

    # Try creating a venv as last resort
    if [ "$PIP_INSTALLED" = false ]; then
        VENV_DIR="$SCRIPT_DIR/.venv"
        echo "Trying to create virtual environment in $VENV_DIR ..."
        if "$PYTHON" -m venv "$VENV_DIR" 2>/dev/null; then
            "$VENV_DIR/bin/pip" install --quiet openpyxl
            PYTHON="$VENV_DIR/bin/python"
            echo "Virtual environment created and openpyxl installed."
            PIP_INSTALLED=true
        fi
    fi

    if [ "$PIP_INSTALLED" = false ]; then
        echo "Error: Could not install openpyxl."
        echo "Install manually: $PYTHON -m pip install openpyxl"
        exit 1
    fi
fi

echo "All dependencies OK. Starting app..."
echo ""
exec "$PYTHON" "$SCRIPT_DIR/main.py"
