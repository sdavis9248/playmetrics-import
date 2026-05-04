@echo off
echo ============================================
echo   PlayMetricsImport — Project Setup
echo ============================================
echo.

REM Check Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python is not installed or not in PATH.
    echo Download from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Found Python:
python --version
echo.

REM Create virtual environment
if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
    echo   Created venv\
) else (
    echo Virtual environment already exists.
)

REM Activate
call venv\Scripts\activate.bat

REM Install dependencies
echo.
echo Installing dependencies...
pip install -r requirements-dev.txt
echo.

REM Create test_data folder
if not exist test_data (
    mkdir test_data
    echo Created test_data\ folder — place your .xlsx files here for testing.
)

echo.
echo ============================================
echo   Setup complete!
echo ============================================
echo.
echo   Next steps:
echo     1. Open this folder in VS Code:  code .
echo     2. VS Code will detect the venv automatically
echo     3. Place playerUpload.xlsx and playerApplications.xlsx
echo        in test_data\ for testing
echo     4. Press F5 to run/debug, or:
echo        python playmetrics_import.py
echo     5. To build the .exe:  build.bat
echo.
pause
