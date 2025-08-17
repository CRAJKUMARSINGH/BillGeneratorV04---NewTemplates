@echo off
echo ====================================
echo  Infrastructure Billing System
echo  Windows 11 Deployment Script
echo ====================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from https://python.org
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo Python is installed. Checking version...
python --version

REM Check if pip is available
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: pip is not available
    echo Please ensure pip is installed with Python
    pause
    exit /b 1
)

echo Installing required packages...
pip install streamlit pandas openpyxl jinja2 reportlab python-docx PyPDF2 beautifulsoup4 lxml

if %errorlevel% neq 0 (
    echo ERROR: Failed to install packages
    pause
    exit /b 1
)

echo.
echo All packages installed successfully!
echo.
echo Starting Infrastructure Billing System...
echo.
echo The application will open in your default web browser.
echo Use Ctrl+C to stop the application.
echo.

REM Start the Streamlit application
streamlit run app.py --server.port 8501 --server.headless false

if %errorlevel% neq 0 (
    echo ERROR: Failed to start the application
    pause
    exit /b 1
)

pause