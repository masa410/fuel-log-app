@echo off
cd /d "%~dp0"
pip install streamlit pytesseract pillow pandas plotly openpyxl > pip_log.txt 2>&1
python -c "import streamlit, pytesseract, pandas, plotly; from PIL import Image; import openpyxl; print('IMPORTS_OK')" > check_imports.txt 2>&1
where tesseract > check_tesseract.txt 2>&1
for /f "delims=" %%i in ('python -c "from local_settings import EXCEL_PATH; print(EXCEL_PATH)"') do set XLSM_PATH=%%i
python import_from_excel.py "%XLSM_PATH%" > import_log.txt 2>&1
echo ALL_DONE > done.txt
