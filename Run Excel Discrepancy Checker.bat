@echo off
cd /d "%~dp0"

REM Check for required packages
python -c "import streamlit, pandas, openpyxl" 2>NUL
if %errorlevel% neq 0 (
    echo Installing required Python packages...
    pip install -r requirements.txt
)

start "" python -m streamlit run app.py 