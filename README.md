# XLSX to JSON Converter

此工具用於將 Excel (.xlsx) 檔案轉換為 JSON 格式，專門設計用於設備參數配置。

## 系統需求

- Python 3.13 或更高版本
- 必要的 Python 套件：
  - pandas
  - openpyxl

## 安裝

1. 複製專案
2. 安裝相依套件：
   ```bash
   pip install -r requirements.txt
   ```

## 使用方式

### 命令格式
```bash
python3 "py_Path" "xlsx_A_Path" "xlsx_B_Path" "DeviceName" "Output_NewPath"
```

### 參數說明
- `py_Path`: Python 執行檔完整路徑
- `xlsx_A_Path`: Excel A 檔案路徑（規格檔 & 保護檔 & 保護回復檔 & 資訊檔一覽表）
- `xlsx_B_Path`: Excel B 檔案路徑（控制檔一覽表）
- `DeviceName`: 設備名稱
- `Output_NewPath`: 輸出目錄路徑

### 命令範例
```bash
python3 "C:/Devin/LTRp_XlsxToJSON/git/xlsx_to_json.py" "C:/Devin/LTRp_XlsxToJSON/git/example/excel_A.xlsx" "C:/Devin/LTRp_XlsxToJSON/git/example/Excel_B.xlsx" "TPT-B3HCR2020A" "C:/Devin/LTRp_XlsxToJSON/git/test_output"
```

## 檔案限制

### Excel A（規格檔 & 保護檔 & 保護回復檔 & 資訊檔一覽表）
- 工作表名稱必須為：
  - "LTRp 規格檔一覽表"
  - "LTRp 保護檔一覽表"
  - "LTRp 保護回復檔一覽表"
  - "LTRp 資訊檔一覽表"

### Excel B（控制檔一覽表）
- 工作表名稱必須為："工作表1"

### 重要注意事項
1. 兩個 Excel 檔案的匯入位置不能互換
2. 每個工作表中的參數名稱需要從 B7 開始
3. 參數名稱中間不能有空值

## 輸出結果
- 在指定的輸出目錄下建立以設備名稱命名的資料夾
- 產生以下 JSON 檔案：
  - DeviceINFO.json
  - CalibrationParemeter.json
  - FWSpecification.json
  - FWControl.json
  - OutputProtection.json
  - ProtectReplyFile.json

## 檔案結構
```
xlsxTojson/
├── xlsx_to_json.py             # 主程式
├── src/
│   ├── excel_processor.py      # Excel 處理模組
│   └── gui.py                  # GUI 介面
└── docs/
    ├── README.md               # 說明文件
    ├── 使用前請先閱讀此份文件.txt    # 使用說明
    └── 匯入檔案限制.txt          # 檔案限制說明
```

## GUI 使用範例
- Python 檔案路徑：`C:/Devin/LTRp_XlsxToJSON/git/xlsx_to_json.py`
- Excel A 路徑：`C:/Devin/LTRp_XlsxToJSON/git/example/excel_A.xlsx`
- Excel B 路徑：`C:/Devin/LTRp_XlsxToJSON/git/example/Excel_B.xlsx`
- 設備名稱：`TPT-B1R4040A`
- 輸出路徑：`C:/Devin/LTRp_XlsxToJSON/git/test_output`
