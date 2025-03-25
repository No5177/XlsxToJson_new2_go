# XlsxToJson 轉換工具

將 Excel 檔案（XLSX）轉換為 JSON 格式的工具，提供圖形化界面操作。

## 功能特點

- 支持批量轉換 Excel 檔案到 JSON 格式
- 自定義輸出 JSON 結構
- 直觀的圖形化界面
- 支持規則檔案作為轉換模板

## 系統需求

- Windows 操作系統
- 無需安裝 Python（獨立執行檔）

## 安裝方法

### 方法一：直接使用可執行檔

從發布頁面下載最新的 `XlsxToJson.exe` 檔案，無需安裝即可使用。

### 方法二：從源碼安裝與構建

1. 確保已安裝 Python 3.8 或更高版本
2. 克隆或下載本代碼庫
3. 安裝依賴：

```bash
pip install -r requirements.txt
```

4. 運行程式：

```bash
python src/gui.py
```

5. 構建獨立執行檔：

```bash
# 使用批處理檔案構建
install.bat
```

構建完成後，可執行檔將位於 `dist` 目錄中。

## 使用說明

1. 啟動程式
2. 選擇來源 Excel 檔案
3. 選擇輸出 JSON 檔案的目錄
4. 設定轉換參數或選擇規則檔案
5. 點擊「開始轉換」按鈕

## 主要檔案說明

- `src/gui.py` - 圖形化界面主程式
- `xlsx_to_json.py` - Excel 到 JSON 的轉換核心模組
- `install.bat` - 構建獨立執行檔的批處理檔案
- `requirements.txt` - Python 依賴庫清單

## 依賴庫

- pandas (>= 2.0.0)
- openpyxl (>= 3.1.0)
- numpy (== 2.2.1)
- tkinter (Python 標準庫) 