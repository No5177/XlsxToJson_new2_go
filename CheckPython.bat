@echo off
if not exist "%~dp0python\python.exe" (
    echo [ERROR] Python installation failed! Please check.
    pause
    exit /b 1
)
exit /b 0
