@echo off
echo Checking Python installation...

where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Error: Python is not installed or not in the PATH.
    echo Please install Python and try again.
    pause
    exit /b
)

echo Checking required packages...
python -c "import flask, pandas, openpyxl" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Installing required packages...
    python -m pip install -r requirements.txt
)

echo Starting Excel Data Cleaner application...
echo Once started, open a web browser and go to http://127.0.0.1:5000
python app.py
pause 