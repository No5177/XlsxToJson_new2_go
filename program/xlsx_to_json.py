import os
import re
import sys
import json
import argparse
import pandas as pd
from decimal import Decimal
from typing import Dict, Any, Tuple, List
from src.excel_processor import ExcelProcessor

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
            
        # First try src directory for spec files
        src_dir = os.path.join(base_path, 'src')
        src_path = os.path.join(src_dir, relative_path)
        if os.path.exists(src_path):
            return src_path

        # Then try with the full relative path
        full_path = os.path.join(base_path, relative_path)
        if os.path.exists(full_path):
            return full_path

        # Then try with just the basename in the root directory
        base_name_path = os.path.join(base_path, os.path.basename(relative_path))
        if os.path.exists(base_name_path):
            return base_name_path

        # If not found in base paths, try current directory
        current_dir_path = os.path.join(os.getcwd(), relative_path)
        if os.path.exists(current_dir_path):
            return current_dir_path

        # If still not found, try script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        script_dir_path = os.path.join(script_dir, relative_path)
        if os.path.exists(script_dir_path):
            return script_dir_path

        raise FileNotFoundError(
            f"Resource not found: {relative_path}\n"
            f"Tried paths:\n"
            f"- {full_path}\n"
            f"- {base_name_path}\n"
            f"- {current_dir_path}\n"
            f"- {script_dir_path}"
        )
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
        processor = ExcelProcessor()
        
        # 載入 Excel 檔案
        excel_data = {
            'A': pd.ExcelFile(excel_a_path),
            'B': pd.ExcelFile(excel_b_path)
        }
        
        # 確保輸出路徑存在
        if output_path is None:
            output_path = os.path.dirname(excel_a_path)
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

        for json_name in json_files:
            try:
                # 特殊處理 CalibrationParemeter.json
                if json_name == "CalibrationParemeter":
                    parameters = {
                        "Gcvadc": 1,
                        "Ocvadc": 0,
                        "Gdvadc": 1,
                        "Odvadc": 0,
                        "Gccadc": 1,
                        "Occadc": 0,
                        "Gdcadc": 1,
                        "Odcadc": 0,
                        "Grvadc": 1,
                        "Orvadc": 0,
                        "Gcvdac": 1,
                        "Ocvdac": 0,
                        "Gdvdac": 1,
                        "Odvdac": 0,
                        "Gccdac": 1,
                        "Occdac": 0,
                        "Gdcdac": 1,
                        "Odcdac": 0
                    }
                else:
                    parameters = processor.process_by_rule(
                        json_name,
                        excel_data,
                        device_name
                    )
                
                # 檢查 JSON 內容
                is_valid, error_msg = check_json_content(json_name, parameters)
                if not is_valid:
                    error_messages.append(error_msg)
                    print(f"{json_name}.json NG")
                    continue
                
                # 特殊處理 DeviceINFO 的 SeriesNumber
                if json_name == "DeviceINFO" and "SeriesNumber" in parameters:
                    raw_series = str(parameters["SeriesNumber"])
                    series_parts = [p.strip() for p in raw_series.split(',')]
                    parameters["SeriesNumber"] = "".join(series_parts)
                
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
                error_messages.append(f"處理 {json_name} 時發生錯誤: {str(e)}")
                print(f"{json_name}.json NG")
        
        if error_messages:
            raise ValueError("\n".join(error_messages))
                
    except Exception as e:
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
