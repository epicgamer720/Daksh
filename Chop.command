#!/bin/bash
cd "$(dirname "$0")"
if ! python3 -c 'import tkinter' 2>/dev/null; then
  echo "Python 3 with tkinter is required. Install Python from https://python.org (the python.org installer includes tkinter)."
  read -p "Press Enter to close..."
  exit 1
fi
# Own venv so pip install works on Homebrew/system Python too (PEP 668)
[ -d .venv ] || python3 -m venv .venv || { read -p "Could not create venv. Press Enter..."; exit 1; }
./.venv/bin/python -c 'import openpyxl, tkinterdnd2, static_ffmpeg' 2>/dev/null \
  || ./.venv/bin/python -m pip install -q -r requirements.txt \
  || { read -p "Dependency install failed. Press Enter..."; exit 1; }
./.venv/bin/python chopper.py || read -p "Press Enter to close..."
