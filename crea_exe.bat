@echo off
REM Crea MercatiReport.exe (da lanciare una volta sola, su un PC Windows con Python installato)
cd /d "%~dp0"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name MercatiReport ^
  --collect-all yfinance --collect-all curl_cffi --collect-data pptx ^
  --hidden-import openpyxl mercati_report.py
echo.
echo Fatto. L'eseguibile si trova in: %~dp0dist\MercatiReport.exe
pause
