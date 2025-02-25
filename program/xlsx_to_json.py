import os
import re
import sys
import json
import argparse
import pandas as pd
import warnings
from decimal import Decimal
from typing import Dict, Any, Tuple, List
from src.excel_processor import ExcelProcessor

# 在程式開始時禁用特定警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # Get the base path for resources
        if getattr(sys, 'frozen', False):
            # Running in PyInstaller bundle
            base_path = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
        else:
            # Running in development
            base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            
        # Convert relative_path to absolute path
        abs_path = os.path.abspath(relative_path)
        if os.path.exists(abs_path):
            return abs_path
            
        # Try with the base path
        base_path_file = os.path.join(base_path, relative_path)
        if os.path.exists(base_path_file):
            return base_path_file
            
        raise FileNotFoundError(f"Resource not found: {relative_path}")
    except Exception as e:
        print(f"Error while locating resource {relative_path}: {str(e)}")
        raise

def validate_device_name(name: str) -> Tuple[bool, str]:
    """Validate device name format."""
    if not name:
        return False, "設備名稱不能為空"
    
    pattern = r'^TPT-[0-9A-Z]+$'
    if not re.match(pattern, name):
        return False, "無效的設備名稱格式。預期格式: TPT-[數字/字母]"
    
    # Basic format validation passed
    return True, ""

class CustomJSONEncoder(json.JSONEncoder):
    def default(self, obj):
        return super().default(obj)
    
    def iterencode(self, obj, _one_shot=False):
        for chunk in super().iterencode(obj, _one_shot=False):
            if chunk.startswith('1e-'):  # 處理科學記號
                # 將科學記號轉換為小數形式
                num = float(chunk)
                chunk = f"{num:.10f}".rstrip('0').rstrip('.')
            yield chunk

def generate_json_file(output_dir: str, json_name: str, parameters: dict) -> str:
    """Generate JSON file based on parameters from Excel"""
    output_path = os.path.join(output_dir, json_name)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(parameters, f, indent=2, ensure_ascii=False, cls=CustomJSONEncoder)
    
    return output_path

def check_json_content(json_name: str, parameters: dict) -> Tuple[bool, str]:
    """檢查 JSON 內容是否有效
    
    Returns:
        Tuple[bool, str]: (是否有效, 錯誤訊息)
    """
    # 檢查是否為空字典或只有一個空值
    if not parameters or (len(parameters) == 1 and list(parameters.values())[0] is None):
        return False, f"{json_name}.json 產生失敗，請檢查excel"
    return True, ""

def check_and_fix_scientific_notation(json_path: str) -> None:
    """檢查 JSON 檔案中的科學符號並修正"""
    try:
        # 讀取 JSON 檔案
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 檢查是否需要修正
        need_fix = False
        for value in data.values():
            if isinstance(value, str) and ('e' in value.lower() or 'E' in value):
                need_fix = True
                break
        
        # 如果需要修正，處理後重新寫入
        #if need_fix:
        #    processed_data = process_parameters(data)
        #    with open(json_path, 'w', encoding='utf-8') as f:
        #        json.dump(processed_data, f, indent=2, ensure_ascii=False)
        #    print(f"Fixed scientific notation in {os.path.basename(json_path)}")
    except Exception as e:
        print(f"Error checking/fixing scientific notation in {json_path}: {str(e)}")

def convert_xlsx_to_json(excel_a_path: str, excel_b_path: str, device_name: str, output_path: str | None = None):
    try:
        # 確保使用絕對路徑
        excel_a_path = os.path.abspath(excel_a_path)
        excel_b_path = os.path.abspath(excel_b_path)
        if output_path:
            output_path = os.path.abspath(output_path)
        else:
            output_path = os.path.dirname(excel_a_path)
            
        # 添加詳細的錯誤輸出
        print(f"處理檔案：\nExcel A: {excel_a_path}\nExcel B: {excel_b_path}\n設備名稱: {device_name}\n輸出路徑: {output_path}")
            
        processor = ExcelProcessor()
        
        # 檢查 Excel A 的工作表名稱
        required_sheets_a = [
            "LTRp 規格檔一覽表",
            "LTRp 保護檔一覽表",
            "LTRp 保護回復檔一覽表",
            "LTRp 資訊檔一覽表"
        ]
        
        # 載入 Excel 檔案並驗證工作表
        try:
            # 使用 pandas 直接讀取 Excel 檔案
            excel_sheets_a = {}
            excel_sheets_b = {}
            
            # 讀取 Excel A 的所有工作表
            for sheet_name in required_sheets_a:
                try:
                    df = pd.read_excel(
                        excel_a_path,
                        sheet_name=sheet_name,
                        engine='openpyxl',
                        header=None  # 不使用標題行
                    )
                    excel_sheets_a[sheet_name] = df
                except Exception as e:
                    raise ValueError(f"無法讀取 Excel A 的工作表 {sheet_name}: {str(e)}")
            
            # 讀取 Excel B 的工作表1
            try:
                df = pd.read_excel(
                    excel_b_path,
                    sheet_name="工作表1",
                    engine='openpyxl',
                    header=None  # 不使用標題行
                )
                excel_sheets_b["工作表1"] = df
            except Exception as e:
                raise ValueError(f"無法讀取 Excel B 的工作表1: {str(e)}")
            
            excel_data = {
                'A': excel_sheets_a,
                'B': excel_sheets_b
            }
            
        except Exception as e:
            raise ValueError(f"讀取 Excel 檔案失敗：{str(e)}")
        
        # 驗證設備名稱格式
        is_valid, error_msg = validate_device_name(device_name)
        if not is_valid:
            raise ValueError(error_msg)
        
        # 確保輸出路徑存在
        os.makedirs(output_path, exist_ok=True)
        
        json_files = [
            "DeviceINFO",
            "CalibrationParemeter",
            "FWSpecification",
            "FWControl",
            "OutputProtection",
            "ProtectReplyFile"
        ]

        error_messages = []
        
        # 定義各 JSON 檔案的處理規則
        json_rules = {
            "DeviceINFO": {
                "sheet_name": "LTRp 資訊檔一覽表",
                "excel_file": "A"
            },
            "FWSpecification": {
                "sheet_name": "LTRp 規格檔一覽表",
                "excel_file": "A"
            },
            "FWControl": {
                "sheet_name": "工作表1",
                "excel_file": "B"
            },
            "OutputProtection": {
                "sheet_name": "LTRp 保護檔一覽表",
                "excel_file": "A"
            },
            "ProtectReplyFile": {
                "sheet_name": "LTRp 保護回復檔一覽表",
                "excel_file": "A"
            }
        }
        
        for json_name in json_files:
            try:
                print(f"正在處理 {json_name}...")
                
                if json_name == "CalibrationParemeter":
                    parameters = {
                        "Gcvadc": 1, "Ocvadc": 0,
                        "Gdvadc": 1, "Odvadc": 0,
                        "Gccadc": 1, "Occadc": 0,
                        "Gdcadc": 1, "Odcadc": 0,
                        "Grvadc": 1, "Orvadc": 0,
                        "Gcvdac": 1, "Ocvdac": 0,
                        "Gdvdac": 1, "Odvdac": 0,
                        "Gccdac": 1, "Occdac": 0,
                        "Gdcdac": 1, "Odcdac": 0
                    }
                else:
                    # 根據規則處理其他 JSON 檔案
                    rule = json_rules.get(json_name)
                    if not rule:
                        raise ValueError(f"找不到 {json_name} 的處理規則")
                    
                    # 獲取對應的工作表數據
                    excel_sheets = excel_data[rule["excel_file"]]
                    sheet_name = rule["sheet_name"]
                    df = excel_sheets[sheet_name]
                    
                    # 從第4行獲取設備型號列表
                    device_list = df.iloc[3]  # 第4行 (索引3)
                    
                    # 找到設備型號所在的列
                    device_col = None
                    for col in range(df.shape[1]):
                        if str(device_list.iloc[col]).strip() == device_name:
                            device_col = col
                            break
                    
                    if device_col is None:
                        raise ValueError(f"在工作表 {sheet_name} 中找不到設備型號 {device_name}")
                    
                    # 使用有序字典來保持參數順序
                    from collections import OrderedDict
                    parameters = OrderedDict()
                    
                    # 處理所有參數
                    for row_idx in range(6, df.shape[0]):  # 從第6行開始
                        row = df.iloc[row_idx]
                        param_name = str(row.iloc[1]).strip()  # 參數名稱在第2列
                        param_value = row.iloc[device_col]  # 參數值在找到的設備型號列
                        
                        if not param_name or pd.isna(param_value):
                            continue
                        
                        # 根據數據類型進行適當的處理
                        if isinstance(param_value, (int, float)):
                            parameters[param_name] = param_value
                        else:
                            # 特殊處理 SeriesNumber
                            if json_name == "DeviceINFO" and param_name == "SeriesNumber":
                                series_parts = [p.strip() for p in str(param_value).split(',')]
                                parameters[param_name] = "".join(series_parts)
                            # 處理需要設為空值的參數
                            elif json_name == "DeviceINFO" and param_name.lower() in ["fwversion", "manufacturedate", "calibrationdate"]:
                                parameters[param_name] = ""
                            else:
                                parameters[param_name] = str(param_value).strip()
                    
                    # 確保 DeviceINFO 中的必要參數存在
                    if json_name == "DeviceINFO":
                        required_empty_params = ["ManufactureDate", "CalibrationDate"]
                        for param_name in required_empty_params:
                            if param_name not in parameters:
                                parameters[param_name] = ""
                
                # 檢查參數是否為空
                if not parameters:
                    raise ValueError(f"{json_name} 的參數為空")
                
                # 檢查 JSON 內容
                is_valid, error_msg = check_json_content(json_name, parameters)
                if not is_valid:
                    error_messages.append(error_msg)
                    print(f"{json_name}.json NG: {error_msg}")
                    continue
                
                # 設定輸出目錄
                if json_name in ["OutputProtection", "ProtectReplyFile"]:
                    file_output_path = os.path.join(output_path, device_name, "Protect")
                else:
                    file_output_path = os.path.join(output_path, device_name)
                
                os.makedirs(file_output_path, exist_ok=True)
                
                # 產生 JSON 檔案
                json_path = generate_json_file(file_output_path, f"{json_name}.json", parameters)
                print(f"{json_name}.json OK")
                
                check_and_fix_scientific_notation(json_path)
                
            except Exception as e:
                error_msg = f"處理 {json_name} 時發生錯誤: {str(e)}"
                error_messages.append(error_msg)
                print(error_msg)
                continue  # 繼續處理下一個檔案
        
        if error_messages:
            raise ValueError("\n".join(error_messages))
                
    except Exception as e:
        print(f"轉換過程發生錯誤: {str(e)}")
        raise ValueError(f"轉換過程發生錯誤: {str(e)}")

def main():
    # Clear any leftover global state
    global excel_a, excel_b_df, device_cols
    excel_a = None
    excel_b_df = None
    device_cols = None
    
    if len(sys.argv) not in [4, 5]:
        print('使用方式: python3 "完整的執行檔路徑" "Excel檔案A路徑.xlsx" "Excel檔案B路徑.xlsx" "設備名稱" ["輸出目錄"]')
        print('例如: python3 "C:/LTRp_XlsxToJSON/git/xlsx_to_json.py" "C:/LTRp_XlsxToJSON/git/excel_A.xlsx" "C:/LTRp_XlsxToJSON/git/excel_B.xlsx" "TPT-B4R1010A" ["C:/Output"]')
        sys.exit(1)
    
    try:
        excel_a_path = os.path.abspath(sys.argv[1])
        excel_b_path = os.path.abspath(sys.argv[2])
        device_name = sys.argv[3]
        # Get output directory if provided
        output_dir = None
        if len(sys.argv) > 4:
            output_dir = sys.argv[4]
            if not os.path.exists(output_dir):
                print(f"警告: 輸出目錄 '{output_dir}' 不存在，將自動建立")
            output_dir = os.path.abspath(output_dir)
        
        # Validate Excel files
        for path, name in [(excel_a_path, 'A'), (excel_b_path, 'B')]:
            if not os.path.exists(path):
                print(f"錯誤: 找不到Excel檔案{name} '{path}'")
                print("請確認Excel檔案的完整路徑是否正確")
                sys.exit(1)
                
            if not path.lower().endswith('.xlsx'):
                print(f"錯誤: 檔案{name}必須是Excel檔案 (.xlsx)")
                sys.exit(1)
            
            
        convert_xlsx_to_json(excel_a_path, excel_b_path, device_name, output_dir)
        
    except Exception as e:
        print(f"錯誤: {str(e)}")
        sys.exit(1)

# Global variables that need to be cleared between runs
excel_a = None
excel_b_df = None
device_cols = None

if __name__ == "__main__":
    main()
