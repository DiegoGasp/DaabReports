@echo off
setlocal enabledelayedexpansion

rem Change to the directory of this script so relative paths work
cd /d "%~dp0"

rem Determine which Python command is available
set "PYTHON_CMD="
where python >nul 2>&1 && set "PYTHON_CMD=python"
if not defined PYTHON_CMD (
    where py >nul 2>&1 && set "PYTHON_CMD=py -3"
)

if not defined PYTHON_CMD (
    echo Python interpreter not found. Please install Python or add it to PATH.
    pause
    exit /b 1
)

%PYTHON_CMD% "%~dp0ParseXml.py" --stream-debug
set "ERR=%ERRORLEVEL%"

if %ERR% neq 0 (
    echo.
    echo The parser encountered an error. Review the messages above and debug.txt for details.
) else (
    echo.
    echo Parsing complete. Output saved to navisworks_views_comments.csv and debug.txt.
)

echo.
pause
