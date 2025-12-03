@echo off
chcp 65001
echo 開始構建獨立執行檔...


if not exist "src\gui.py" (
  echo Error：Not find src\gui.py 
  pause
  exit /b 1
)

if not exist "xlsx_to_json.py" (
  echo Error：Not find xlsx_to_json.py 
  pause
  exit /b 1
)

if not exist "icon.ico" (
  echo Alarm：Not find icon.ico , use defut icon
)


if not exist "dist" (
  mkdir dist
)


if not exist "dist\data" (
  mkdir dist\data
)

nuitka src/gui.py 
--standalone ^
--onefile ^
--enable-plugin=tk-inter ^
--windows-icon-from-ico=icon.ico ^
--windows-disable-console ^
--windows-company-name="ThinkPower" ^
--windows-product-name="ThinkRP Spec Excel to JSON Converter" ^
--windows-file-description="ThinkRP Spec Excel to JSON Converter" ^
--windows-product-version="1.0.0.4" ^
--windows-file-version="1.0.0.4" ^
--include-data-file=icon.ico=icon.ico ^
--include-module=xlsx_to_json ^
--include-data-dir=icon_image=icon_image ^
--follow-imports ^
--output-dir=dist ^
--output-filename=ThinkRP Spec Excel to JSON Converter.exe ^
--include-package=pandas ^
--include-package=openpyxl ^
--include-package=numpy ^
--include-package=pathlib ^


REM --- 新增錯誤檢查邏輯 ---
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo 🔴 Nuitka 編譯失敗，錯誤碼：%ERRORLEVEL%。
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ✅ Nuitka 編譯成功！
REM --- 正常情況下繼續執行，直到最後的暫停 ---

pause