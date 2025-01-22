@echo off
echo 開始安裝 LTRp_XlsxToJSON...

REM 檢查 Python 是否已安裝
python --version > nul 2>&1
if errorlevel 1 (
    echo Python 未安裝，正在進行無聲安裝...
    start /wait python\python-3.13.1-amd64 /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
)

REM 安裝必要套件
echo 安裝必要套件...
pip install --no-index --find-links=packages pandas openpyxl numpy

REM 建立程式目錄
set INSTALL_DIR=%USERPROFILE%\LTRp_XlsxToJSON
echo 建立程式目錄: %INSTALL_DIR%
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

REM 複製程式檔案
echo 複製程式檔案...
xcopy /E /I /Y ..\program\* "%INSTALL_DIR%"

echo.
echo 安裝完成！
echo 程式已安裝到: %INSTALL_DIR%
echo 請參考 docs/使用前請先閱讀此份文件.txt 了解使用方式
pause 