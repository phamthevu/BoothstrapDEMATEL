@echo off
cd /d %~dp0

echo Installing dependencies...
pip install -r requirements.txt

echo Running app...
streamlit run app.py

pause