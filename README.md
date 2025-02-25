# Excel 參數檔轉換工具

此工具用於將 Excel (.xlsx) 檔案轉換為 JSON 格式的參數檔，專門用於設備參數配置管理。

## 系統需求

- Python 3.11 或更高版本
- 必要的 Python 套件：
  - pandas (用於 Excel 檔案處理)
  - openpyxl (用於 .xlsx 檔案支援)

## 安裝與執行

1. 下載並解壓縮專案檔案
2. 執行 `run2.bat` 啟動圖形介面

> 注意：首次執行時會自動安裝必要的 Python 套件

## 使用說明

### 圖形介面操作步驟

1. 程式啟動後會自動載入 xlsx_to_json.py 的路徑
2. 點選「瀏覽」按鈕選擇以下檔案：
   - Excel 檔案 A（規格檔/保護檔/保護回復檔/資訊檔）
   - Excel 檔案 B（控制檔）
3. 輸入設備型號（例：TPT-B30R10100A）
4. 選擇輸出目錄
5. 點選「所有參數檔生成」開始轉換

### Excel 檔案格式要求

#### Excel 檔案 A
必須包含以下工作表：
- "LTRp 規格檔一覽表"
- "LTRp 保護檔一覽表"
- "LTRp 保護回復檔一覽表"
- "LTRp 資訊檔一覽表"

#### Excel 檔案 B
- 工作表名稱必須為："工作表1"

### 注意事項
1. Excel A 與 B 檔案不可互換
2. 參數名稱需從 B7 儲存格開始
3. 參數名稱之間不可有空值
4. 設備型號格式必須為：TPT-[數字及大寫英文字母]

## 輸出結果

程式會在指定的輸出目錄下建立以設備型號命名的資料夾，並產生以下 JSON 檔案：

```
設備型號/
├── DeviceINFO.json          # 設備資訊
├── CalibrationParemeter.json # 校正參數
├── FWSpecification.json     # 韌體規格
├── FWControl.json          # 韌體控制
└── Protect/
    ├── OutputProtection.json    # 輸出保護
    └── ProtectReplyFile.json    # 保護回覆
```

## 專案結構
```
XlsxToJson/
├── run2.bat                # 執行檔
├── program/
│   ├── xlsx_to_json.py     # 轉換核心程式
│   ├── requirements.txt    # 套件需求檔
│   └── src/
│       └── gui.py         # 圖形介面程式
└── README.md              # 說明文件
```

## 錯誤排除

如果遇到執行錯誤，請檢查：
1. Python 是否正確安裝
2. Excel 檔案格式是否符合要求
3. 檔案路徑是否包含特殊字元
4. 輸出目錄是否有寫入權限

## 支援與回報問題

如發現任何問題或需要協助，請：
1. 檢查命令執行狀態視窗的錯誤訊息
2. 確認 Excel 檔案格式是否正確
3. 聯繫技術支援人員
