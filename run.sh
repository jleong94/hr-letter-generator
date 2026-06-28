#!/bin/bash
# HR Letter Generator launcher for macOS / Linux.
set -u
cd "$(dirname "$0")"

echo "============================================"
echo "   HR Letter Generator"
echo "============================================"
echo

find_python() {
  if command -v python3 >/dev/null 2>&1; then
    PY="python3"
  elif command -v python >/dev/null 2>&1; then
    PY="python"
  else
    PY=""
  fi
}

find_python
if [ -z "$PY" ]; then
  echo "Python 3 was not found. Attempting to install it..."
  if command -v brew >/dev/null 2>&1; then
    brew install python
    find_python
  else
    echo
    echo "[!] Homebrew is not installed, so Python cannot be installed automatically."
    echo "    Option 1: install Homebrew from https://brew.sh then run this again."
    echo "    Option 2: install Python 3 from https://www.python.org/downloads/"
    echo
    read -r -p "Press Enter to close..."
    exit 1
  fi
fi

if [ -z "$PY" ]; then
  echo "[!] Python still not found. Please reopen this launcher and try again."
  read -r -p "Press Enter to close..."
  exit 1
fi

echo "Using Python: $($PY --version 2>&1)"
echo

if [ ! -x ".venv/bin/python" ]; then
  echo "Creating virtual environment..."
  "$PY" -m venv .venv || { echo "[!] Failed to create the virtual environment."; read -r -p "Press Enter to close..."; exit 1; }
fi
VENVPY=".venv/bin/python"

if [ ! -f ".venv/.deps_ok" ]; then
  echo "Installing required packages, please wait..."
  "$VENVPY" -m pip install --upgrade pip
  "$VENVPY" -m pip install -r requirements.txt || { echo "[!] Failed to install packages. Check your internet connection."; read -r -p "Press Enter to close..."; exit 1; }
  echo "ok" > ".venv/.deps_ok"
fi

echo
"$VENVPY" generate_letter.py
echo
read -r -p "Press Enter to close..."
