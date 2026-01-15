@echo off
REM DGN File Version Sorter
REM Drag and drop a folder onto this batch file, or run it in a folder with DGN files

echo ============================================================
echo DGN File Version Sorter
echo ============================================================

if "%~1"=="" (
    echo No folder specified, using current directory...
    python "%~dp0sort_dgn_by_version.py" "%cd%"
) else (
    echo Sorting files in: %~1
    python "%~dp0sort_dgn_by_version.py" "%~1"
)

echo.
pause
