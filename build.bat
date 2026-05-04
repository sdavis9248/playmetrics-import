@echo off
echo ============================================
echo   Building PlayMetricsImport.exe
echo ============================================
echo.

REM Activate virtual environment
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
) else (
    echo ERROR: Virtual environment not found.
    echo Run setup.bat first to create it.
    pause
    exit /b 1
)

REM Build the exe
echo Building single-file executable...
pyinstaller --onefile --name PlayMetricsImport playmetrics_import.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Build failed.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Build complete!
echo ============================================
echo.
echo   Output:  dist\PlayMetricsImport.exe
echo   Size:    
for %%A in (dist\PlayMetricsImport.exe) do echo            %%~zA bytes
echo.
echo   To distribute, copy these files together:
echo     dist\PlayMetricsImport.exe
echo     README.txt
echo.

REM Copy README.txt to dist for easy packaging
if exist README.txt (
    copy README.txt dist\README.txt >nul
    echo   README.txt copied to dist\ folder.
    echo.
)

pause
