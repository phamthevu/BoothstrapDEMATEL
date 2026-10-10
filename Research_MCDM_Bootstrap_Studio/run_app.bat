@echo off
setlocal
cd /d "%~dp0"
python run_app.py
if errorlevel 1 (
  echo.
  echo App failed to start. Check that Python and requirements are installed.
  pause
)
