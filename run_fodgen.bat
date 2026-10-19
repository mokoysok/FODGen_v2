@echo off
setlocal

set "PROJECT_DIR=%~dp0"
set "PYTHON_EXE=%PROJECT_DIR%.venv\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
    echo.
    echo Python 3.13 virtual environment not found.
    echo Run: py -3.13 -m venv .venv
    echo Then: .venv\Scripts\python.exe -m pip install -r requirements.txt
    pause
    exit /b 1
)

"%PYTHON_EXE%" "%PROJECT_DIR%main.py"
if errorlevel 1 pause
