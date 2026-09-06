@echo off
cd /d "%~dp0"
pip install streamlit pytesseract pillow pandas plotly openpyxl > pip_log.txt 2>&1
python -c "import streamlit, pytesseract, pandas, plotly; from PIL import Image; import openpyxl; print('IMPORTS_OK')" > check_imports.txt 2>&1
where tesseract > check_tesseract.txt 2>&1
python import_from_excel.py "C:\Users\今枝真紀\OneDrive\デスクトップ\真紀フォルダー\家計管理関連\ソリオ燃費早見表.xlsm" > import_log.txt 2>&1
echo ALL_DONE > done.txt
