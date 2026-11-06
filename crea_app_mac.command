#!/bin/bash
# Crea MercatiReport.app per Mac (lanciare una volta sola, su un Mac con Python di python.org)
cd "$(dirname "$0")"
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt pyinstaller
python3 -m PyInstaller --noconfirm --windowed --name MercatiReport \
  --collect-all yfinance --collect-all curl_cffi --collect-data pptx \
  --hidden-import openpyxl mercati_report.py
ditto -c -k --keepParent dist/MercatiReport.app MercatiReport_mac.zip
echo
echo "Fatto. App: dist/MercatiReport.app  (zip da inviare: MercatiReport_mac.zip)"
