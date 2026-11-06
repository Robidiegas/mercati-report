@echo off
REM Avvio diretto con Python (alternativa all'exe)
cd /d "%~dp0"
python -m pip install -q -r requirements.txt
start "" pythonw mercati_report.py
