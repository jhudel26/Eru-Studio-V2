@echo off
title EruStudio Launcher
echo Starting EruStudio...
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from https://python.org
    echo.
    pause
    exit /b 1
)

REM Check if requirements are installed
echo Checking dependencies...
pip show openpyxl >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install required packages
        echo Please check your internet connection and try again
        echo.
        pause
        exit /b 1
    )
)

REM Launch the application
echo Launching EruStudio...
python main.py

REM Check if the application exited with an error
if errorlevel 1 (
    echo.
    echo EruStudio encountered an error. Please check the error message above.
    echo.
    pause
) 