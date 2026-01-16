@echo off
setlocal enabledelayedexpansion

echo ============================================================
echo DWG Cleaner - FAST MODE
echo ============================================================
echo.
echo This version uses optimized timings for faster processing.
echo Use standard clean_dwg.bat if you encounter errors.
echo.

:: Get the directory where the batch file is located
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

:: Output folder
set "OUTPUT_DIR=%SCRIPT_DIR%\CLEAN"
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

echo Input folder:  %SCRIPT_DIR%
echo Output folder: %OUTPUT_DIR%
echo.

:: Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/
    pause
    exit /b 1
)

:: Check if pywin32 is installed
python -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
    echo Installing required package: pywin32...
    pip install pywin32
    if errorlevel 1 (
        echo ERROR: Failed to install pywin32.
        pause
        exit /b 1
    )
)

:: Run the Python script
echo Starting FAST MODE cleanup...
echo.
python "%SCRIPT_DIR%\clean_dwg_fast.py" "%SCRIPT_DIR%"

echo.
echo ============================================================
echo Cleanup complete! Check the CLEAN folder.
echo ============================================================
pause
