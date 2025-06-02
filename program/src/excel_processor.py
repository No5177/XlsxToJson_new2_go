import os
import re
import pandas as pd
from typing import Dict, Any, Tuple, Optional

# Parameter processing configuration
# Note: Removed series specifications as per requirements
# JSON structure is maintained without value restrictions

class ExcelProcessor:
    def read_excel(self, file_path: str) -> Optional[pd.DataFrame]:
        """Read Excel file and return DataFrame."""
        try:
            return pd.read_excel(file_path)
        except Exception as e:
            print(f"Error reading Excel file: {str(e)}")
            return None

    def validate_device_name(self, name: str, df: pd.DataFrame) -> Tuple[bool, str]:
        """Validate device name format and existence in Excel."""
        if not name:
            return False, "設備名稱不能為空"
        
        pattern = r'^TPT-[0-9A-Z]+$'
        if not re.match(pattern, name):
            return False, "無效的設備名稱格式。預期格式: TPT-[數字/字母]"
        
        # Check if device exists in Excel
        device_row = df.iloc[2]  # Row 3 (0-based index 2)
        if name not in device_row.values:
            return False, f"找不到設備 '{name}'"
        
        # Basic format validation passed
        return True, ""

    def get_device_column(self, df: pd.DataFrame, device_name: str) -> Optional[int]:
        """Get the column index for the specified device."""
        device_col = None
        for row_idx in range(len(df)):
            row_values = df.iloc[row_idx].values
            
            # Check each cell in the row
            for col_idx, val in enumerate(row_values):
                if pd.notnull(val) and str(val).strip() == device_name:
                    device_col = col_idx
                    break
            
            if device_col is not None:
                break
            
        # Verify the column contains valid data
        if device_col is not None:
            valid_values = [
                val for val in df.iloc[:, device_col].values 
                if pd.notnull(val) and str(val).strip() and str(val).strip() != device_name
            ]
            if not valid_values:
                return None
            
        return device_col

    def extract_parameters(self, df: pd.DataFrame, device_col: int) -> Dict[str, Any]:
        """Extract parameters starting from B6 until first empty B cell."""
        parameters = {}
        
        # Start from B6 (index 5) and process rows until first empty B cell
        row_index = 5  # B6
        while row_index < len(df):
            try:
                # Get parameter name from column B (index 1)
                param_name = df.iloc[row_index, 1]
                
                # Break if we hit an empty cell in column B
                if pd.isna(param_name) or str(param_name).strip() == "":
                    break
                    
                param_value = df.iloc[row_index, device_col]
                
                # Convert parameter name to string and strip
                param_name = str(param_name).strip()
                
                # Skip special rows
                if param_name in ["Name", "Device", "紅色字體屬於使用者輸入欄位"]:
                    row_index += 1
                    continue
                
                # Clean parameter name if it has a numeric prefix
                json_key = param_name
                if '.' in param_name:
                    parts = param_name.split('.')
                    if len(parts) == 2 and parts[1].strip():
                        json_key = parts[1].strip()
                
                # 處理數值，保持原始格式
                if pd.notnull(param_value):
                    if isinstance(param_value, float):
                        str_val = f"{param_value:.10f}".rstrip('0').rstrip('.')
                        parameters[json_key] = int(str_val) if str_val.isdigit() else float(str_val)
                    elif isinstance(param_value, str):
                        try:
                            # 嘗試將字串轉換為浮點數，以處理小數和負號
                            converted_val = float(param_value)
                            # 如果轉換後的浮點數等於其整數形式，則轉換為整數
                            if converted_val == int(converted_val):
                                parameters[json_key] = int(converted_val)
                            else:
                                parameters[json_key] = converted_val
                        except ValueError:
                            # 如果字串無法轉換為有效數字，則保留為字串
                            parameters[json_key] = param_value
                    else:
                        parameters[json_key] = param_value
                else:
                    parameters[json_key] = 0 if json_key in ["UVP", "CISETmin", "DISETmin"] else None
                
            except Exception:
                pass
            
            row_index += 1
            
        return parameters

    def process_by_rule(self, json_name: str, excel_data: Dict[str, pd.ExcelFile], device_name: str) -> Dict[str, Any]:
        """Process parameters based on specific rules defined in text files."""
        parameters = {}
        rule_file_path = os.path.join("Generate files_spec_rule", f"{json_name}_rule.txt")
        
        if not os.path.exists(rule_file_path):
            print(f"Warning: Rule file not found: {rule_file_path}")
            return parameters

        with open(rule_file_path, 'r', encoding='utf-8') as f:
            rules = f.readlines()

        if json_name == "DeviceINFO":
            # 特定處理 DeviceINFO.json
            df_a = excel_data['A'].parse("LTRp 資訊檔一覽表")
            device_col_a = self.get_device_column(df_a, device_name)
            if device_col_a is not None:
                parameters = self.extract_parameters(df_a, device_col_a)
                parameters["Fwversion"] = ""
                parameters["ManufactureDate"] = ""
                parameters["CalibrationDate"] = ""
                if "SeriesNumber" in parameters:
                    parameters["SeriesNumber"] = "".join(parameters["SeriesNumber"].split(','))

        # 在此處添加其他 JSON 檔案的處理邏輯
        elif json_name == "CalibrationParemeter":
            df_a = excel_data['A'].parse("輸出固定項目與固定值")
            device_col_a = self.get_device_column(df_a, device_name)
            if device_col_a is not None:
                parameters = self.extract_parameters(df_a, device_col_a)
        elif json_name == "FWSpecification":
            df_a = excel_data['A'].parse("LTRp 規格檔一覽表")
            device_col_a = self.get_device_column(df_a, device_name)
            if device_col_a is not None:
                parameters = self.extract_parameters(df_a, device_col_a)
        elif json_name == "FWControl":
            df_b = excel_data['B'].parse("工作表1")
            device_col_b = self.get_device_column(df_b, device_name)
            if device_col_b is not None:
                parameters = self.extract_parameters(df_b, device_col_b)
        elif json_name == "OutputProtection":
            df_a = excel_data['A'].parse("LTRp 保護檔一覽表")
            device_col_a = self.get_device_column(df_a, device_name)
            if device_col_a is not None:
                parameters = self.extract_parameters(df_a, device_col_a)
        elif json_name == "ProtectReplyFile":
            df_a = excel_data['A'].parse("LTRp 保護回復檔一覽表")
            device_col_a = self.get_device_column(df_a, device_name)
            if device_col_a is not None:
                parameters = self.extract_parameters(df_a, device_col_a)

        return parameters

    # Removed parameter validation method as per requirements
    # Parameters are now processed without range restrictions
