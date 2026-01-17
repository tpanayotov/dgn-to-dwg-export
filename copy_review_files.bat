@echo off
setlocal enabledelayedexpansion

echo ============================================================
echo Copy Files Marked for Review to REDO folder
echo ============================================================
echo.

:: Get the directory where the batch file is located
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

:: Look for the report in CLEAN subfolder
set "REPORT_PATH=%SCRIPT_DIR%\CLEAN\clean_dwg_report.csv"

if not exist "%REPORT_PATH%" (
    echo ERROR: Report file not found at:
    echo   %REPORT_PATH%
    echo.
    echo Please ensure clean_dwg_report.csv exists in the CLEAN folder.
    pause
    exit /b 1
)

:: Create REDO folder
set "REDO_DIR=%SCRIPT_DIR%\REDO"
if not exist "%REDO_DIR%" mkdir "%REDO_DIR%"

echo Source report: %REPORT_PATH%
echo Output folder: %REDO_DIR%
echo.

:: Count files to copy
set "COUNT=0"
set "COPIED=0"

:: Skip header line, find lines with "review" status
for /f "skip=1 tokens=1,2 delims=," %%a in ('type "%REPORT_PATH%"') do (
    set "FILENAME=%%~a"
    set "STATUS=%%~b"

    if /i "!STATUS!"=="review" (
        set /a COUNT+=1

        :: Source file is in the same folder as the batch file
        set "SOURCE_FILE=%SCRIPT_DIR%\!FILENAME!"

        if exist "!SOURCE_FILE!" (
            echo Copying: !FILENAME!
            copy "!SOURCE_FILE!" "%REDO_DIR%\" >nul
            if !errorlevel! equ 0 (
                set /a COPIED+=1
            ) else (
                echo   WARNING: Failed to copy !FILENAME!
            )
        ) else (
            echo   WARNING: Source file not found: !FILENAME!
        )
    )
)

echo.
echo ============================================================
echo Done! Found !COUNT! files marked for review.
echo Copied !COPIED! files to REDO folder.
echo ============================================================
pause
