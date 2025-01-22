import os
import json
import pytest
import pandas as pd
from pathlib import Path
from unittest.mock import patch, MagicMock
from xlsx_to_json import convert_xlsx_to_json, generate_json_file, validate_device_name

@pytest.fixture
def sample_excel_files(tmp_path):
    """Create sample Excel files for testing."""
    # Create Excel A with required sheets
    excel_a_path = tmp_path / "test_excel_a.xlsx"
    with pd.ExcelWriter(excel_a_path) as writer:
        # 規格檔一覽表
        # Create DataFrames with correct structure
        spec_df = pd.DataFrame([
            ['1', '1. VSETmax', 20000],
            ['2', '2. VSETmin', 400],
            ['Device', 'Name', 'TPT-TEST123'],
            ['3', '3. CISETmax', 10000]
        ], columns=['No.', '紅色字體屬於使用者輸入欄位', 'TPT-TEST123'])
        spec_df.to_excel(writer, sheet_name='LTRp 規格檔一覽表', index=False)
        
        # 保護回復檔一覽表
        protect_df = pd.DataFrame([
            ['1', '1. OVP', 21000],
            ['2', '2. UVP', 0],
            ['Device', 'Name', 'TPT-TEST123'],
            ['3', '3. OCP', 10500],
            ['4', '4. OTP', 85],
            ['5', '5. OPP', 2000],
            ['6', '6. UCP', 0],
            ['7', '7. RVP', 500],
            ['8', '8. COCP', 11000],
            ['9', '9. DOCP', 11000],
            ['10', '10. BOCP', 11000],
            ['11', '11. POCP', 11000],
            ['12', '12. FOCP', 11000],
            ['13', '13. AlarmDelay', 3]
        ], columns=['No.', '紅色字體屬於使用者輸入欄位', 'TPT-TEST123'])
        protect_df.to_excel(writer, sheet_name='LTRp 保護回復檔一覽表', index=False)
        
        # 資訊檔一覽表
        info_df = pd.DataFrame([
            ['1', '1. VSETmax', 20000],
            ['2', '2. VSETmin', 400],
            ['Device', 'Name', 'TPT-TEST123'],
            ['3', '3. CISETmax', 10000],
            ['4', '4. CISETmin', 0],
            ['5', '5. DISETmax', 10000],
            ['6', '6. DISETmin', 0],
            ['7', '7. PSETmax', 2000],
            ['8', '8. PSETmin', 0],
            ['9', '9. Fwversion', ''],
            ['10', '10. ManufactureDate', ''],
            ['11', '11. CalibrationDate', '']
        ], columns=['No.', '紅色字體屬於使用者輸入欄位', 'TPT-TEST123'])
        info_df.to_excel(writer, sheet_name='LTRp 資訊檔一覽表', index=False)
    
    # Create Excel B
    excel_b_path = tmp_path / "test_excel_b.xlsx"
    control_df = pd.DataFrame([
        ['1', 'ControlMode', 1],
        ['2', 'ControlValue', 100],
        ['Device', 'Name', 'TPT-TEST123'],
        ['3', 'ControlTime', 1000]
    ], columns=['Unnamed: 0', '紅色字體屬於使用者輸入欄位', 'TPT-TEST123'])
    control_df.to_excel(excel_b_path, index=False)
    
    return str(excel_a_path), str(excel_b_path)

def test_generate_json_file(tmp_path):
    """Test JSON file generation from template."""
    output_dir = str(tmp_path)
    template_name = "DeviceINFO.json"  # Will be converted to _spec.json internally
    parameters = {
        "DeviceName": "TPT-TEST123",
        "VSETmax": 20000,
        "VSETmin": 400,
        "Fwversion": "",
        "ManufactureDate": "",
        "CalibrationDate": ""
    }
    
    # No need to copy template - generate_json_file will find it in src/
    
    output_path = generate_json_file(output_dir, template_name, parameters)
    
    assert os.path.exists(output_path)
    with open(output_path, 'r') as f:
        data = json.load(f)
        assert data["DeviceName"] == "TPT-TEST123"
        assert data["VSETmax"] == 20000
        assert data["VSETmin"] == 400
        assert data["Fwversion"] == ""
        assert data["ManufactureDate"] == ""
        assert data["CalibrationDate"] == ""

def test_convert_xlsx_to_json_directory_structure(tmp_path, sample_excel_files):
    """Test the creation of correct directory structure and files."""
    excel_a_path, excel_b_path = sample_excel_files
    device_name = "TPT-TEST123"
    output_path = str(tmp_path / "output")
    
    convert_xlsx_to_json(excel_a_path, excel_b_path, device_name, output_path)
    
    # Check directory structure
    device_dir = os.path.join(output_path, device_name)
    protect_dir = os.path.join(device_dir, "Protect")
    assert os.path.exists(device_dir)
    assert os.path.exists(protect_dir)
    
    # Check all JSON files exist
    assert os.path.exists(os.path.join(device_dir, "FWSpecification.json"))
    assert os.path.exists(os.path.join(device_dir, "DeviceINFO.json"))
    assert os.path.exists(os.path.join(device_dir, "CalibrationParemeter.json"))
    assert os.path.exists(os.path.join(device_dir, "FWControl.json"))
    assert os.path.exists(os.path.join(protect_dir, "OutputProtection.json"))
    assert os.path.exists(os.path.join(protect_dir, "ProtectReplyFile.json"))

def test_calibration_parameter_values(tmp_path, sample_excel_files):
    """Test that CalibrationParemeter.json uses fixed values."""
    excel_a_path, excel_b_path = sample_excel_files
    device_name = "TPT-TEST123"
    output_path = str(tmp_path / "output")
    
    convert_xlsx_to_json(excel_a_path, excel_b_path, device_name, output_path)
    
    calib_path = os.path.join(output_path, device_name, "CalibrationParemeter.json")
    with open(calib_path, 'r') as f:
        data = json.load(f)
        # Check gain values are 1
        assert data["GV"] == 1
        assert data["GI"] == 1
        assert data["GD"] == 1
        assert data["GP"] == 1
        assert data["GT"] == 1
        assert data["GmAh"] == 1
        assert data["GWh"] == 1
        # Check offset values are 0
        assert data["OV"] == 0
        assert data["OI"] == 0
        assert data["OD"] == 0
        assert data["OP"] == 0
        assert data["OT"] == 0
        assert data["OmAh"] == 0
        assert data["OWh"] == 0

def test_device_info_special_fields(tmp_path, sample_excel_files):
    """Test that DeviceINFO.json has correct special empty fields."""
    excel_a_path, excel_b_path = sample_excel_files
    device_name = "TPT-TEST123"
    output_path = str(tmp_path / "output")
    
    convert_xlsx_to_json(excel_a_path, excel_b_path, device_name, output_path)
    
    info_path = os.path.join(output_path, device_name, "DeviceINFO.json")
    with open(info_path, 'r') as f:
        data = json.load(f)
        assert data["Fwversion"] == ""
        assert data["ManufactureDate"] == ""
        assert data["CalibrationDate"] == ""
        assert data["DeviceName"] == device_name

def test_device_name_validation():
    """Test device name validation with various formats."""
    valid_names = [
        "TPT-123",
        "TPT-ABC123",
        "TPT-B123",  # Still valid but B not required
        "TPT-TEST123"
    ]
    invalid_names = [
        "",  # Empty
        "TPT123",  # Missing hyphen
        "ABC-123",  # Wrong prefix
        "TPT-123!",  # Invalid character
        "TPT--123"  # Double hyphen
    ]
    
    for name in valid_names:
        valid, _ = validate_device_name(name)
        assert valid, f"Name should be valid: {name}"
    
    for name in invalid_names:
        valid, _ = validate_device_name(name)
        assert not valid, f"Name should be invalid: {name}"

def test_excel_worksheet_validation(tmp_path):
    """Test validation of required worksheets in Excel A."""
    # Create Excel with missing worksheet
    excel_path = tmp_path / "invalid_excel_a.xlsx"
    with pd.ExcelWriter(excel_path) as writer:
        pd.DataFrame().to_excel(writer, sheet_name='LTRp 規格檔一覽表')
        # Missing 保護回復檔一覽表 and 資訊檔一覽表
    
    with pytest.raises(ValueError) as exc_info:
        convert_xlsx_to_json(str(excel_path), "dummy.xlsx", "TPT-TEST123", str(tmp_path))
    assert "Excel檔案A缺少工作表: LTRp 保護回復檔一覽表" in str(exc_info.value)
    assert "Excel檔案A缺少工作表" in str(exc_info.value)

def test_parameter_value_preservation(tmp_path, sample_excel_files):
    """Test that parameter values from Excel are preserved without modification."""
    excel_a_path, excel_b_path = sample_excel_files
    device_name = "TPT-TEST123"
    output_path = str(tmp_path / "output")
    
    convert_xlsx_to_json(excel_a_path, excel_b_path, device_name, output_path)
    
    # Check FWSpecification.json values match Excel
    spec_path = os.path.join(output_path, device_name, "FWSpecification.json")
    with open(spec_path, 'r') as f:
        data = json.load(f)
        assert data["VSETmax"] == 20000
        assert data["VSETmin"] == 400
        assert data["CISETmax"] == 10000
    
    # Check protection values match Excel
    protect_path = os.path.join(output_path, device_name, "Protect", "OutputProtection.json")
    with open(protect_path, 'r') as f:
        data = json.load(f)
        assert data["OVP"] == 21000
        assert data["UVP"] == 0
        assert data["OCP"] == 10500
