@echo off
echo Starting program...
echo.

:: Check if Python is installed
python --version > nul 2>&1
if errorlevel 1 (
    echo Python is not installed! Please install Python first.
    pause
    exit /b 1
)

:: Check if file exists
if not exist src\gui.py (
    echo src\gui.py file not found!
    pause
    exit /b 1
)

:: Check if virtual environment exists
if exist venv (
    :: Activate virtual environment
    call venv\Scripts\activate
)

:: Run program without waiting
start /b "" python src\gui.py

:: Exit the batch file immediately
exit

:: Deactivate if virtual environment was used
if exist venv (
    deactivate
)

pause 