@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 exit /b 1
pyinstaller --noconfirm --clean --windowed --name ResearchMCDMBootstrap run_app.py
if errorlevel 1 exit /b 1
echo Build complete: dist\ResearchMCDMBootstrap\ResearchMCDMBootstrap.exe
pause
