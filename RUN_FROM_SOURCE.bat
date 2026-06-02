@echo off
cd /d "%~dp0"
echo Starting Mail Exporter from source (no .exe required)...
echo Requires: Python 3.8+ and Microsoft Outlook on this PC.
echo.
python eml_to_pst_converter.py
if errorlevel 1 (
  echo.
  echo If Python is missing, install from https://www.python.org/downloads/
  echo Then run: python -m pip install -r requirements.txt
  pause
)
