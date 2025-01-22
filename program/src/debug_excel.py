import pandas as pd
import sys

def analyze_excel(file_path: str, device_name: str | None = None):
    print(f"Analyzing Excel file: {file_path}")
    if device_name:
        print(f"Looking for device: {device_name}\n")
    
    # Read Excel file
    df = pd.read_excel(file_path)
    
    # Print column names
    print("Column names:")
    print(df.columns.tolist())
    print("\n")
    
    # Get all device names from row 2
    print("All device names in row 2:")
    device_row = df.iloc[2]  # Row 3 (0-based index 2)
    device_names = device_row[2:].dropna().tolist()  # Skip first two columns
    for i, name in enumerate(device_names, 1):
        print(f"{i}. {name}")
    
    if not device_name:
        return device_names
    
    # Find device column
    device_cols = [i for i, val in enumerate(device_row.values) if val == device_name]
    if device_cols:
        device_col = device_cols[0]
        print(f"\nFound device {device_name} in column {device_col}")
    else:
        print(f"\nDevice {device_name} not found")
        return device_names
    
    # Print parameter rows
    print("\nParameter rows:")
    param_rows = df[df['Unnamed: 0'].astype(str).str.match(r'^[1-9]\d*$|^8[AB]$', na=False)]
    for _, row in param_rows.iterrows():
        param_num = str(row['Unnamed: 0'])
        param_name = row['紅色字體屬於使用者輸入欄位'].strip()
        param_value = row.iloc[device_col]
        print(f"No.{param_num}: {param_name} = {param_value}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python debug_excel.py <excel_file> [device_name]")
        sys.exit(1)
    
    device_name = sys.argv[2] if len(sys.argv) > 2 else None
    analyze_excel(sys.argv[1], device_name)
