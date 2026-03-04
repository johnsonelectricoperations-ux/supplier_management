@echo off
chcp 65001 > nul
cd /d %~dp0

if not exist venv (
    echo [Setup] Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo [Error] Python not found. Please install Python first.
        pause
        exit /b 1
    )
)

call venv\Scripts\activate.bat

echo [Setup] Installing packages...
pip install -q -r requirements.txt

if not exist pdfs mkdir pdfs
if not exist db mkdir db

echo.
echo ============================================
echo   Inspection Management System
echo   URL: http://10.80.101.200:5002
echo ============================================
echo.

venv\Scripts\uvicorn.exe main:app --host 0.0.0.0 --port 5002 --workers 1 --access-log --log-level info
pause
