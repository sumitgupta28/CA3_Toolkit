@echo off
setlocal enabledelayedexpansion

echo ============================================
echo  CA3 Toolkit - Environment Setup
echo ============================================
echo.

:: Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo        Download from https://www.python.org/downloads/
    echo        Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYTHON_VERSION=%%v
echo Found: %PYTHON_VERSION%
echo.

:: Create virtual environment
if exist .venv (
    echo Virtual environment already exists, skipping creation.
) else (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo Done.
)
echo.

:: Activate virtual environment
echo Activating virtual environment...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment.
    pause
    exit /b 1
)
echo Done.
echo.

:: Upgrade pip silently
echo Upgrading pip...
python -m pip install --upgrade pip --quiet
echo Done.
echo.

:: Install dependencies
echo Installing dependencies from requirements.txt...
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)
echo.

echo ============================================
echo  Setup complete!
echo ============================================
echo.
echo To activate the environment in future sessions, run:
echo     .venv\Scripts\activate
echo.
echo Workflow:
echo   1. python extract_to_excel.py
echo   2. Open marks.xlsx and fill in marks
echo   3. python fill_marks_and_export.py
echo.

cmd /k
