@echo off
echo 開始構建獨立執行檔...

REM 檢查目錄結構
if not exist "src\gui.py" (
  echo 錯誤：找不到 src\gui.py 檔案
  pause
  exit /b 1
)

if not exist "xlsx_to_json.py" (
  echo 錯誤：找不到 xlsx_to_json.py 檔案
  pause
  exit /b 1
)

if not exist "icon.ico" (
  echo 警告：找不到 icon.ico 檔案，將使用默認圖標
)

REM 創建輸出目錄
if not exist "dist" (
  mkdir dist
)

REM 創建data目錄用於複製樣例Excel文件
if not exist "dist\data" (
  mkdir dist\data
)

REM 執行 Nuitka 打包命令
REM --windows-disable-console 已註解掉
nuitka src/gui.py ^
  --standalone ^
  --onefile ^
  --enable-plugin=tk-inter ^
  --windows-icon-from-ico=icon.ico ^
  --windows-disable-console ^
  --windows-company-name="ThinkPower" ^
  --windows-product-name="ThinkRP Spec Excel to JSON Converter" ^
  --windows-file-description="ThinkRP Spec Excel to JSON Converter" ^
  --windows-product-version="1.0.0.3" ^
  --windows-file-version="1.0.0.3" ^
  --include-data-file=icon.ico=icon.ico ^
  --include-module=xlsx_to_json ^
  --include-data-dir=icon_image=icon_image ^
  --include-data-file=xlsx_to_json.py=xlsx_to_json.py ^
  --follow-imports ^
  --output-dir=dist ^
  --output-filename=ThinkRP Spec Excel to JSON Converter.exe ^
  --nofollow-import-to=test_gui,test_json_generation,debug_excel,src.test_gui,src.test_json_generation,src.debug_excel ^
  --include-package=pandas ^
  --include-package=openpyxl ^
  --include-package=numpy ^
  --include-package=pathlib

REM 複製範例Excel檔案到dist/data目錄（若有的話）
if exist "example_excel\*.xlsx" (
  copy "example_excel\*.xlsx" "dist\data\"
)

echo.
echo 構建完成！
echo 可執行檔位於 dist\XlsxToJson.exe
echo Excel範例檔案（若有）已複製到 dist\data 目錄
pause
