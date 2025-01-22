# XLSX to JSON Converter Implementation Plan

## 1. User Interface Design
### Main Window
```
+----------------------------------------+
|  XLSX to JSON Converter                |
|----------------------------------------|
|  [ Select XLSX File... ]               |
|  File: [                          ]    |
|                                        |
|  Device Name: [                   ]    |
|                                        |
|  [ Convert ]                           |
|                                        |
|  Status: Ready                         |
+----------------------------------------+
```

### Features
- File selection button and display field
- Device name input field with validation
- Convert button
- Status display area
- Error message display capability

## 2. Core Components

### 2.1 File Processing Module
```python
class ExcelProcessor:
    def read_excel(self, file_path: str) -> pd.DataFrame:
        """Read Excel file and validate basic structure."""
        
    def extract_parameters(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Extract parameters No.1~95 (including 8A, 8B).
        
        Parameter mapping:
        - No.1-20: Voltage settings (VSET*)
        - No.21-40: Current settings (CISET*, DISET*)
        - No.41-60: Power settings (PSET*, CPSET*)
        - No.61-80: Time settings
        - No.81-95: Protection settings
        - No.8A, 8B: Special current parameters
        """
        
    def validate_parameters(self, params: Dict[str, Any]) -> bool:
        """Validate parameters against specification."""
        
    def map_parameters(self, raw_params: Dict[str, Any], device_type: str) -> Dict[str, Any]:
        """Map Excel parameters to JSON structure based on device type."""
```

### 2.1.1 Parameter Extraction Details
1. Required Parameters (No.1~95):
   ```
   1. VSETmax          31. DISETmax         61. Timesetmax
   2. VSETmin          32. DISETmin         62. Timesetmin
   3. VSETreso         33. DISETreso        ...
   ...                 8A. CISETspecial     91. OVPmax
   8. VMEANmax         8B. DISETspecial     92. UVPmax
   ...                 40. CISETACCmin      ...
   ```

2. Parameter Categories:
   - Voltage Parameters (V*)
   - Current Parameters (CI*, DI*)
   - Power Parameters (P*)
   - Time Parameters
   - Protection Parameters
   - Special Parameters (8A, 8B)

### 2.2 Device Name Validator
```python
class DeviceValidator:
    def validate_device_name(self, name: str) -> bool
    def get_device_specs(self, name: str) -> Dict[str, Any]
```

### 2.3 JSON Generator
```python
class JsonGenerator:
    def create_json(self, params: Dict[str, Any], device_name: str) -> bool
    def validate_json(self, json_data: Dict[str, Any]) -> bool
```

## 3. Implementation Flow

1. File Selection
   - Allow .xlsx file selection
   - Validate file format and content
   - Display selected file path

2. Device Name Input
   - Accept device name input
   - Validate against format pattern (TPT-B[number/letter])
   - Check against parameter ranges based on device series

3. Parameter Extraction
   - Read Excel file using pandas
   - Extract parameters No.1~95 (including 8A, 8B)
   - Map parameters to JSON structure
   - Validate parameter values against ranges

4. Output Generation
   - Create device-named folder
   - Generate FWSpecification.json
   - Validate JSON against specification

## 4. Validation Rules Implementation

### 4.1 Device Name Validation

#### 4.1.1 Device Name Format
```python
def validate_device_name_format(name: str) -> bool:
    """Validate device name matches TPT-B[series][specs] format."""
    pattern = r'^TPT-B[0-9A-Z]+$'
    return bool(re.match(pattern, name))
```

#### 4.1.2 Device Series Validation
```python
def validate_device_series(name: str) -> Tuple[bool, str]:
    """Validate device series and get corresponding specifications.
    
    Known series:
    - B1RS: 5V/5A series (VSETmax=5000, CISETmax=5000)
    - B4R: 20V/10A series (VSETmax=20000, CISETmax=10000)
    """
    series_specs = {
        'B1RS': {'VSETmax': 5000, 'CISETmax': 5000},
        'B4R': {'VSETmax': 20000, 'CISETmax': 10000}
    }
    
    series = re.search(r'TPT-B(\d+[A-Z]+)', name)
    if not series:
        return False, "Invalid series format"
        
    series_id = series.group(1)
    if series_id not in series_specs:
        return False, f"Unknown series: {series_id}"
        
    return True, series_id
```

#### 4.1.3 Specification Validation
```python
def validate_against_spec(name: str, params: Dict[str, Any]) -> bool:
    """Validate device parameters against LTRp specification.
    
    Steps:
    1. Extract series from device name
    2. Get series specifications
    3. Validate parameters against series ranges
    4. Check for required parameters (No.1~95, 8A, 8B)
    5. Validate parameter relationships
    """
```

### 4.2 Parameter Validation
```python
def validate_parameters(params: Dict[str, Any]) -> bool:
    rules = {
        'VSETmax': (0, 20000),
        'CISETmax': (0, 10000),
        'PSETmax': (0, 200),
        # ... other parameters
    }
    return all(rules[k][0] <= v <= rules[k][1] for k, v in params.items())
```

## 5. Installation Package

### 5.1 Dependencies
```
requirements.txt:
- pandas
- openpyxl
- tkinter (for UI)
- pytest (for testing)
```

### 5.2 Build Process
1. Create virtual environment
2. Install dependencies
3. Package application using PyInstaller
4. Create installer using NSIS

### 5.3 Directory Structure
```
xlsxTojson/
├── src/
│   ├── __init__.py
│   ├── ui/
│   ├── excel_processor/
│   ├── json_generator/
│   └── validators/
├── tests/
├── requirements.txt
├── setup.py
└── README.md
```

## 6. Testing Strategy

### 6.1 Unit Tests
- Test parameter extraction
- Test device name validation
- Test JSON generation
- Test value range validation

### 6.2 Integration Tests
- Test complete conversion process
- Test UI interaction
- Test file handling
- Test error scenarios

### 6.3 Test Data
- Sample Excel files
- Known device configurations
- Invalid input cases

## 7. Error Handling

### 7.1 User Input Errors
- Invalid file format
- Invalid device name
- Missing required parameters

### 7.2 Processing Errors
- File access errors
- Parameter extraction errors
- JSON generation errors

### 7.3 System Errors
- Disk space issues
- Permission issues
- Memory constraints

## 8. Implementation Timeline

1. Core Components (2 days)
   - Excel processing
   - Parameter validation
   - JSON generation

2. UI Development (1 day)
   - Main window
   - File selection
   - Input validation

3. Testing (1 day)
   - Unit tests
   - Integration tests
   - Error handling

4. Packaging (1 day)
   - Installer creation
   - Documentation
   - Final testing
