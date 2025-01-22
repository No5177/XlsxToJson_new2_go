@echo off
echo Starting installation of LTRp_XlsxToJSON...

REM Check if Python is installed
python --version > nul 2>&1
if errorlevel 1 (
    echo Python is not installed. Starting silent installation...
    start /wait python\python-3.13.1-amd64 /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
)

REM Install required packages
echo Installing required packages...
pip install --no-index --find-links=packages pandas openpyxl numpy

REM Create program directory
set INSTALL_DIR=%USERPROFILE%\LTRp_XlsxToJSON
echo Creating program directory: %INSTALL_DIR%
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

REM Copy program files
echo Copying program files...
xcopy /E /I /Y ..\program\* "%INSTALL_DIR%"

echo.
echo Installation completed!
echo Program installed to: %INSTALL_DIR%
echo Please refer to docs/README.txt for usage instructions
pause 