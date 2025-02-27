@echo off
:: 關閉命令回顯

:: 使用 vbs 腳本來隱藏命令行窗口
echo Creating VBS script...
echo Set WshShell = CreateObject("WScript.Shell") > "%temp%\invisible.vbs"
echo WshShell.Run chr(34) ^& "%~f0" ^& chr(34) ^& " run", 0 >> "%temp%\invisible.vbs"
echo Set WshShell = Nothing >> "%temp%\invisible.vbs"

:: 如果沒有參數，則執行 vbs 腳本並退出
if "%~1"=="" (
    start "" "%temp%\invisible.vbs"
    exit
)

:: 主程序邏輯
:: 檢查是否安裝了 Python
python --version > nul 2>&1
if errorlevel 1 (
    echo Python is not installed! Please install Python first.
    pause
    exit /b 1
)

:: 檢查程序主文件是否存在
if not exist program\src\gui.py (
    echo program\src\gui.py file not found!
    pause
    exit /b 1
)

:: 檢查必要的套件是否已安裝
python -c "import pandas" 2>nul
if errorlevel 1 (
    echo Installing required packages...
    python -m pip install --user pandas openpyxl > nul 2>&1
)

:: 啟動程序
start /wait /b python program\src\gui.py

:: 刪除臨時 vbs 腳本
del "%temp%\invisible.vbs" > nul 2>&1

exit 