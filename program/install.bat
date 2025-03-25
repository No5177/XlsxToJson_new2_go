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

REM 執行 Nuitka 打包命令
nuitka src/gui.py ^
  --standalone ^
  --onefile ^
  --enable-plugin=tk-inter ^
  --windows-disable-console ^
  --windows-icon-from-ico=icon.ico ^
  --include-data-file=icon.ico=icon.ico ^
  --include-module=xlsx_to_json ^
  
  --include-data-dir=icon_image=icon_image ^
  --include-data-file=xlsx_to_json.py=xlsx_to_json.py ^
  --follow-imports ^
  --output-dir=dist ^
  --output-filename=XlsxToJson.exe ^
  --nofollow-import-to=test_gui,test_json_generation,debug_excel,src.test_gui,src.test_json_generation,src.debug_excel ^
  --nofollow-import-to=pandas,numpy,openpyxl ^
  --include-package=pandas ^
  --include-package=openpyxl ^
  --include-package=numpy ^
  --include-package=pathlib

echo.
echo 構建完成！
echo 可執行檔位於 dist\XlsxToJson.exe
pause
