@echo off
cd /d "%~dp0"
py -3 -c "import openpyxl, tkinterdnd2, static_ffmpeg, PIL" 2>nul || py -3 -m pip install -r requirements.txt
py -3 chopper.py
if errorlevel 1 pause
