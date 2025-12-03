# **網站工具需求規格書** 

## **I. 總覽 (Overview)**

本工具旨在根據使用者指定的規格檔（包含規格、保護、資訊等工作表）和控制檔，生成一組包含 6 個 JSON 檔案和 1 個資料夾的 ZIP 壓縮包，用於設備配置。本工具需解決不同工作表間設備型號欄位不一致的問題，並針對特定檔案應用特殊的數據清洗規則。

## **II. A. Presentation Layer（介面層）**

**職責：** 處理使用者互動、顯示數據、介面導航與防呆機制。

| 組件名稱 | 欄位名稱 | 輸入類型/操作 | 說明與狀態行為 |
| :---- | :---- | :---- | :---- |
| **1\. 檔案匯入區** | 指定 規格檔(.xlsx)路徑 | 輸入 (File) **\+ 必填標示(紅色\*)** | 用於選擇核心規格檔案。 未選擇時不可進行後續操作。 |
| **2\. 檔案匯入區** | 指定 控制檔(.xlsx)路徑 | 輸入 (File) **\+ 必填標示(紅色\*)** | 用於選擇控制檔案。 未選擇時不可進行後續操作。 |
| **3\. 檔案匯入區** | 選擇設備輸出型號 | 下拉選單 (Dropdown) \+ "讀取型號"按鈕 | **預設狀態：禁用 (Disabled)，樣式顯示為灰色。** 行為：點擊「讀取型號」並成功解析後，解除禁用狀態，並填充設備型號清單。 |
| **4\. 檔案匯入區** | 指定產出檔案路徑 | 輸入 (String/Path) | *(Web版行為)* 瀏覽器下載行為，路徑由瀏覽器設定決定。 |
| **5\. 操作控制區** | 執行生成檔案 | 按鈕 | 觸發業務邏輯，生成並壓縮檔案。 |
| **6\. 狀態與紀錄區** | 紀錄顯示 | 顯示區塊 (Log View) | 顯示操作進度、檔案驗證狀態、成功訊息或錯誤提示。 |

## **III. B. Application/Business Logic Layer (應用/業務邏輯層)**

**職責：** 定義應用程式執行特定任務的流程，協調數據層和介面層。

### **1\. 檔案輸入與驗證 (File Input & Validation)**

* **規格檔 (specFile) 驗證：** 匯入的檔名必須以 POCB-V100A 規格檔 & 保護檔 & 保護回復檔 & 資訊檔一覽表 為開頭。  
* **控制檔 (ctrlFile) 驗證：** 匯入的檔名必須以 POCB-V100A 控制檔一覽表 為開頭。  
* **檔名範例：** POCB-V100A 控制檔一覽表\_V1.20\_20250721\_...xlsx。

### **2\. XlsxToJson\_RP\_SPEC（業務邏輯核心）**

#### **A. 讀取設備型號功能 (UI 清單用)**

* **目的：** 僅用於填充前端下拉選單供使用者選擇。  
* **來源：** 規格檔 (specFile) 的 "LTRp 規格檔一覽表" 分頁。  
* **邏輯：** 讀取 Row 4，從 Column 6 (F 欄) 開始往右讀取所有非空儲存格。

#### **B. 獨立欄位搜尋邏輯 (CRITICAL \- 修正邏輯)**

* **規則：** 由於相同型號在不同工作表（或不同檔案）中可能位於不同的欄位 (Column)，**禁止**假設所有工作表的型號欄位索引相同。  
* **執行流程：** 在生成每一個 JSON 檔案時，程式必須進入該檔案的對應**來源工作表**，在 **Row 4** 重新搜尋使用者所選的「設備型號」字串，以取得該工作表專屬的 **Column Index**。

#### **C. JSON 數據解析通用邏輯**

* **參數名稱 (Key) 來源：** 位於 Column 2 (B 欄)，從 Row 7 開始向下讀取。  
* **參數數值 (Value) 來源：** 位於步驟 B 找到的對應 Column，從 Row 7 開始向下讀取。  
* **數據類型與精度約束：**  
  * **嚴格保持 Excel 原始類型：** 讀取到的數值必須保持 Excel 儲存格中的原始數據類型（字串或數字）。  
  * **精度保留：** 若為數值，必須完全保留小數位數（例如 Excel 顯示 1.00，JSON 須為 1.00 或確保精度未丟失）。

### **3\. FileGenerator（檔案生成與輸出）**

#### **A. 檔案生成流程與特殊規則**

根據所選設備型號，針對不同目標檔案執行以下處理：

**1\. DeviceINFO.json (特殊規則)**

* **空值處理 (Empty Value Rule)：** 若 Excel 中對應的 Value 為空，**必須保留 Key**，並將 Value 寫入空字串 "" (String format)。(即：不可遺失欄位)。  
* **參數 FWversion：** 強制覆寫為空字串 ""。  
* **參數 SeriesNumber：** 讀取數值後，必須移除所有逗號 ,，並以文字格式儲存。  
  * *範例：* 讀取到 "04,16405,00001" $\\rightarrow$ 寫入 "041640500001"。

**2\. FWControl.json (特殊規則)**

* **空值處理 (Empty Value Rule)：** 若 Excel 中對應的 Value 為空，**必須保留 Key**，並將 Value 寫入空字串 ""。  
  * *說明：* 這是為了解決控制檔內容可能為空導致產出 {} 的問題。

**3\. 其他檔案 (FWSpecification, OutputProtection, etc.)**

* **空值處理：** 若 Excel 中對應的 Value 為空，則**跳過該欄位** (不寫入 JSON Key)。

#### **B. 最終輸出與命名**

* **ZIP 檔名：** 選定設備型號\_SPEC \+ yyyymmddHHMMSS (時間戳記)。  
  * *範例：* TPT-B4R200200B\_SPEC\_20251203160000.zip  
* **ZIP 內容結構：**

\<ZIP 根目錄\>  
├── Protect/ (檔案夾)  
│   ├── OutputProtection.json  
│   └── ProtectReplyFile.json  
├── CalibrationParemeter.json  
├── DeviceINFO.json  
├── FWControl.json  
├── FWSpecification.json  
└── LogFile.json

## **IV. C. GenerateJson.md (檔案生成參考詳細表)**

| JSON 檔案名稱 | 來源檔案 | 來源工作表 (Sheet Name) | 空值處理規則 | 特殊處理備註 |
| :---- | :---- | :---- | :---- | :---- |
| **OutputProtection.json** | 規格檔 | "LTRp 保護檔一覽表" | 跳過欄位 | 無 |
| **ProtectReplyFile.json** | 規格檔 | "LTRp 保護回復檔一覽表" | 跳過欄位 | 無 |
| **DeviceINFO.json** | 規格檔 | "LTRp 資訊檔一覽表" | **保留 Key，填入 ""** | 1\. FWversion 固定為 "" 2\. SeriesNumber 去除逗號 , |
| **FWSpecification.json** | 規格檔 | "LTRp 規格檔一覽表" | 跳過欄位 | 無 |
| **FWControl.json** | 控制檔 | "工作表1" | **保留 Key，填入 ""** | 確保即使 Excel 無值也能產出完整 Key 列表 |
| **CalibrationParemeter.json** | 無 | 無 (固定內容) | N/A | 使用下方定義的常數 |

### **CalibrationParemeter.json 固定內容**

{  
  "Gcvadc": 1, "Ocvadc": 0, "Gdvadc": 1, "Odvadc": 0,  
  "Gccadc": 1, "Occadc": 0, "Gdcadc": 1, "Odcdac": 0,  
  "Grvadc": 1, "Orvadc": 0, "Gcvdac": 1, "Ocvdac": 0,  
  "Gdvdac": 1, "Odvdac": 0, "Gccdac": 1, "Occdac": 0,  
  "Gdcdac": 1, "Odcdac": 0  
}  
