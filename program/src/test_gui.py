import unittest
from unittest.mock import MagicMock, patch, call
import tkinter as tk
import os
import sys
from pathlib import Path

# Add src directory to Python path
sys.path.append(str(Path(__file__).parent))

from gui import XlsxToJsonGUI

class MockStringVar:
    """Simple mock for tkinter's StringVar that supports value tracking"""
    def __init__(self, master=None, value=''):
        self._value = str(value)
        self._callbacks = []
        
    def get(self):
        return self._value
        
    def set(self, value):
        old_value = self._value
        self._value = str(value)
        if old_value != self._value:
            for callback in self._callbacks:
                callback('', '', 'w')  # Match tkinter's callback signature
            
    def trace_add(self, mode, callback):
        if mode == 'write':
            self._callbacks.append(callback)
        
    def __str__(self):
        return self.get()
        
    def __eq__(self, other):
        if isinstance(other, MockStringVar):
            return self.get() == other.get()
        return str(self.get()) == str(other)

class MockWidget:
    def __init__(self, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs
        
    def grid(self, **kwargs):
        self.grid_kwargs = kwargs
        return self

@patch('tkinter.StringVar', MockStringVar)
@patch('tkinter.ttk.Frame', MockWidget)
@patch('tkinter.ttk.Label', MockWidget)
@patch('tkinter.ttk.Entry', MockWidget)
@patch('tkinter.ttk.Button', MockWidget)
@patch('os.path.abspath', lambda x: x)  # Make abspath return input unchanged
class TestXlsxToJsonGUI(unittest.TestCase):
    def setUp(self):
        self.root = MagicMock()
        self.root.title = MagicMock()
        self.mock_showerror = MagicMock()
        self.mock_showinfo = MagicMock()
        self.gui = XlsxToJsonGUI(
            self.root,
            string_var_class=MockStringVar,
            messagebox_showerror=self.mock_showerror,
            messagebox_showinfo=self.mock_showinfo
        )
    
    def test_path_storage(self):
        """Test path storage in StringVar variables"""
        test_py_path = "/path/to/script.py"
        test_xlsx_path = "/path/to/data.xlsx"
        test_device = "TPT-B30R10100A"
        test_output = "/path/to/output"
        
        self.gui.py_Path.set(test_py_path)
        self.gui.xlsx_Path_A.set(test_xlsx_path)
        self.gui.xlsx_Path_B.set(test_xlsx_path)
        self.gui.DeviceName.set(test_device)
        self.gui.Output_Path.set(test_output)
        
        self.assertEqual(self.gui.py_Path.get(), test_py_path)
        self.assertEqual(self.gui.xlsx_Path_A.get(), test_xlsx_path)
        self.assertEqual(self.gui.xlsx_Path_B.get(), test_xlsx_path)
        self.assertEqual(self.gui.DeviceName.get(), test_device)
        self.assertEqual(self.gui.Output_Path.get(), test_output)
    
    def test_py_file_selection(self):
        """Test Python file selection"""
        test_path = "/path/to/script.py"
        with patch('tkinter.filedialog.askopenfilename', return_value=test_path) as mock_dialog:
            self.gui.select_py_file()
            self.assertEqual(self.gui.py_Path.get(), test_path)
            mock_dialog.assert_called_with(
                title="選擇Python檔案",
                filetypes=[("Python files", "*.py")]
            )
    
    def test_xlsx_file_selection(self):
        """Test Excel file selection for both files"""
        test_path = "/path/to/data.xlsx"
        
        # Test Excel A selection
        with patch('tkinter.filedialog.askopenfilename', return_value=test_path) as mock_dialog:
            self.gui.select_xlsx_file_A()
            self.assertEqual(self.gui.xlsx_Path_A.get(), test_path)
            mock_dialog.assert_called_with(
                title="選擇Excel檔案 (規格檔 & 保護檔 & 保護回復檔 & 資訊檔一覽表)",
                filetypes=[("Excel files", "*.xlsx")]
            )
            
        # Test Excel B selection
        with patch('tkinter.filedialog.askopenfilename', return_value=test_path) as mock_dialog:
            self.gui.select_xlsx_file_B()
            self.assertEqual(self.gui.xlsx_Path_B.get(), test_path)
            mock_dialog.assert_called_with(
                title="選擇Excel檔案 (控制檔一覽表)",
                filetypes=[("Excel files", "*.xlsx")]
            )
    
    def test_output_path_selection(self):
        """Test output path selection"""
        test_path = "/path/to/output"
        with patch('tkinter.filedialog.askdirectory', return_value=test_path) as mock_dialog:
            self.gui.select_output_path()
            self.assertEqual(self.gui.Output_Path.get(), test_path)
            mock_dialog.assert_called_with(title="選擇輸出目錄")
    
    def test_generate_fw_spec_success(self):
        """Test successful FWSpecification generation"""
        # Setup test paths
        py_path = "C:\\path\\to\\script.py"
        xlsx_path_a = "C:\\path\\to\\data_a.xlsx"
        xlsx_path_b = "C:\\path\\to\\data_b.xlsx"
        device_name = "TPT-B30R10100A"
        output_path = "C:\\path\\to\\output"
        
        # Set paths in GUI
        self.gui.py_Path.set(py_path)
        self.gui.xlsx_Path_A.set(xlsx_path_a)
        self.gui.xlsx_Path_B.set(xlsx_path_b)
        self.gui.DeviceName.set(device_name)
        self.gui.Output_Path.set(output_path)
        
        # Expected converted paths
        expected_py_path = "C:/path/to/script.py"
        expected_xlsx_path_a = "C:/path/to/data_a.xlsx"
        expected_xlsx_path_b = "C:/path/to/data_b.xlsx"
        expected_output_path = "C:/path/to/output"
        
        with patch('subprocess.run') as mock_run, \
             patch('os.chdir') as mock_chdir, \
             patch('os.path.isfile') as mock_isfile, \
             patch('os.makedirs') as mock_makedirs:
            # Configure mocks
            mock_run.return_value = MagicMock(returncode=0)
            mock_isfile.return_value = True
            
            # Execute generate_fw_spec
            self.gui.generate_fw_spec()
            
            # Verify sequence of operations using call_args_list
            # First, verify makedirs was called
            self.assertEqual(len(mock_makedirs.call_args_list), 1)
            makedirs_call = mock_makedirs.call_args_list[0]
            self.assertEqual(makedirs_call, call(expected_output_path, exist_ok=True))
            
            # Then verify chdir was called with output path before the command
            chdir_calls = mock_chdir.call_args_list
            self.assertGreaterEqual(len(chdir_calls), 2)  # At least two calls: to output dir and back
            self.assertEqual(chdir_calls[0], call(expected_output_path))
            
            # Finally verify command execution
            self.assertEqual(len(mock_run.call_args_list), 1)
            run_call = mock_run.call_args_list[0]
            self.assertEqual(
                run_call,
                call(
                    f'python3 "{expected_py_path}" "{expected_xlsx_path_a}" "{expected_xlsx_path_b}" "{device_name}" "{expected_output_path}"',
                    shell=True,
                    capture_output=True,
                    text=True
                )
            )
            
            # Verify success message
            self.mock_showinfo.assert_called_with("成功", "所有參數檔已成功產生")
    
    def test_validation_empty_fields(self):
        """Test validation with empty fields"""
        # Set all values to empty strings
        self.gui.py_Path.set("")
        self.gui.xlsx_Path_A.set("")
        self.gui.xlsx_Path_B.set("")
        self.gui.DeviceName.set("")
        self.gui.Output_Path.set("")
        
        # Execute generate_fw_spec
        self.gui.generate_fw_spec()
        
        # Verify error message
        self.mock_showerror.assert_called_with("錯誤", "請填寫所有必要欄位")
    
    def test_validation_command_error(self):
        """Test validation when command execution fails"""
        # Set all required fields
        self.gui.py_Path.set("/path/to/script.py")
        self.gui.xlsx_Path_A.set("/path/to/data.xlsx")
        self.gui.xlsx_Path_B.set("/path/to/control.xlsx")
        self.gui.DeviceName.set("TPT-B30R10100A")
        self.gui.Output_Path.set("/path/to/output")
        
        with patch('subprocess.run') as mock_run, \
             patch('os.chdir'), \
             patch('os.path.isfile') as mock_isfile, \
             patch('os.makedirs'):
            # Configure mocks
            mock_run.return_value = MagicMock(returncode=1, stderr="Command failed")
            mock_isfile.return_value = True
            
            # Execute generate_fw_spec
            self.gui.generate_fw_spec()
            
            # Verify error message
            self.mock_showerror.assert_called_with("錯誤", "執行失敗:\nCommand failed")

if __name__ == '__main__':
    unittest.main()
