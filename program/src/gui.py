import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import subprocess

class XlsxToJsonGUI:
    def __init__(self, root, string_var_class=None, messagebox_showerror=None, messagebox_showinfo=None):
        self.root = root
        self.root.title("XLSX to JSON Converter")
        
        # Set string var class (use tkinter's if not provided)
        self._string_var_class = string_var_class if string_var_class is not None else tk.StringVar
        
        # Set messagebox functions (use tkinter's if not provided)
        self._messagebox_showerror = messagebox_showerror if messagebox_showerror is not None else messagebox.showerror
        self._messagebox_showinfo = messagebox_showinfo if messagebox_showinfo is not None else messagebox.showinfo
        
        # Variables to store paths and device name
        self._py_Path = ""
        self._xlsx_Path_A = ""  # Main Excel file
        self._xlsx_Path_B = ""  # Control file
        self._device_name = ""
        self._output_path = ""
        
        # GUI variables
        self.py_Path = self._string_var_class(value=self._py_Path)
        self.xlsx_Path_A = self._string_var_class(value=self._xlsx_Path_A)
        self.xlsx_Path_B = self._string_var_class(value=self._xlsx_Path_B)
        self.DeviceName = self._string_var_class(value=self._device_name)
        self.Output_Path = self._string_var_class(value=self._output_path)
        
        # Set trace callbacks
        self.py_Path.trace_add("write", lambda *args: setattr(self, '_py_Path', self.py_Path.get()))
        self.xlsx_Path_A.trace_add("write", lambda *args: setattr(self, '_xlsx_Path_A', self.xlsx_Path_A.get()))
        self.xlsx_Path_B.trace_add("write", lambda *args: setattr(self, '_xlsx_Path_B', self.xlsx_Path_B.get()))
        self.DeviceName.trace_add("write", lambda *args: setattr(self, '_device_name', self.DeviceName.get()))
        self.Output_Path.trace_add("write", lambda *args: setattr(self, '_output_path', self.Output_Path.get()))
        
        # Set icon
        try:
            icon_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'icon_image', 'icon_517.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Warning: Could not load icon: {e}")
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        # 設定固定寬度
        label_width = 30
        entry_width = 50
        entry_padx = 5  # 統一的 Entry padding
        
        # Python file selection
        ttk.Label(main_frame, text="指定python檔", width=label_width).grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.py_Path, width=entry_width).grid(row=0, column=1, padx=entry_padx)
        ttk.Button(main_frame, text="瀏覽", command=self.select_py_file).grid(row=0, column=2, padx=5)
        
        # Excel file A selection (Main configuration file)
        ttk.Label(main_frame, text="指定Excel檔案(規格檔 & 保護檔 & 保護回復檔 & 資訊檔一覽表)", width=label_width).grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.xlsx_Path_A, width=entry_width).grid(row=1, column=1, padx=entry_padx)
        ttk.Button(main_frame, text="瀏覽", command=self.select_xlsx_file_A).grid(row=1, column=2, padx=5)
        
        # Excel file B selection (Control file)
        ttk.Label(main_frame, text="指定Excel檔案(控制檔一覽表)", width=label_width).grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.xlsx_Path_B, width=entry_width).grid(row=2, column=1, padx=entry_padx)
        ttk.Button(main_frame, text="瀏覽", command=self.select_xlsx_file_B).grid(row=2, column=2, padx=5)
        
        # Device name input
        ttk.Label(main_frame, text="指定設備型號", width=label_width).grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.DeviceName, width=entry_width).grid(row=3, column=1, padx=5)
        ttk.Label(main_frame, text="例: TPT-B30R10100A", foreground="gray").grid(row=3, column=2, padx=5)
        
        # Output path selection
        ttk.Label(main_frame, text="指定產出檔案路徑", width=label_width).grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.Output_Path, width=entry_width).grid(row=4, column=1, padx=entry_padx)
        ttk.Button(main_frame, text="瀏覽", command=self.select_output_path).grid(row=4, column=2, padx=5)
        
        # Command output display
        ttk.Label(main_frame, text="命令執行狀態", width=label_width).grid(row=5, column=0, sticky=tk.W, pady=5)
        self.cmd_output = tk.Text(main_frame, height=10, width=70)
        self.cmd_output.grid(row=5, column=1, columnspan=2, padx=entry_padx, pady=5)
        
        # Generate button
        ttk.Button(main_frame, text="所有參數檔生成", command=self.generate_fw_spec).grid(row=6, column=1, pady=20)
        
    def select_py_file(self):
        filename = filedialog.askopenfilename(
            title="選擇Python檔案",
            filetypes=[("Python files", "*.py")]
        )
        if filename:
            self.py_Path.set(filename)
    
    def select_xlsx_file_A(self):
        filename = filedialog.askopenfilename(
            title="選擇Excel檔案 (規格檔 & 保護檔 & 保護回復檔 & 資訊檔一覽表)",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if filename:
            self.xlsx_Path_A.set(filename)
            
    def select_xlsx_file_B(self):
        filename = filedialog.askopenfilename(
            title="選擇Excel檔案 (控制檔一覽表)",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if filename:
            self.xlsx_Path_B.set(filename)
    
    def select_output_path(self):
        directory = filedialog.askdirectory(title="選擇輸出目錄")
        if directory:
            self.Output_Path.set(directory)
    
    def generate_fw_spec(self):
        # Get values from StringVar objects and strip whitespace
        py_path = str(self.py_Path.get()).strip()
        xlsx_path_a = str(self.xlsx_Path_A.get()).strip()
        xlsx_path_b = str(self.xlsx_Path_B.get()).strip()
        device_name = str(self.DeviceName.get()).strip()
        output_path = str(self.Output_Path.get()).strip()
        
        # Validate inputs
        if not all([py_path, xlsx_path_a, xlsx_path_b, device_name, output_path]):
            self._messagebox_showerror("錯誤", "請填寫所有必要欄位")
            return
        
        # Convert paths to use forward slashes
        py_NewPath = py_path.replace("\\", "/")
        xlsx_NewPath_A = xlsx_path_a.replace("\\", "/")
        xlsx_NewPath_B = xlsx_path_b.replace("\\", "/")
        Output_NewPath = output_path.replace("\\", "/")
        
        try:
            # Create output directory first
            os.makedirs(Output_NewPath, exist_ok=True)
            
            # Validate file existence
            if not os.path.isfile(py_NewPath):
                self._messagebox_showerror("錯誤", f"找不到Python檔案: {py_NewPath}")
                return
                
            if not os.path.isfile(xlsx_NewPath_A):
                self._messagebox_showerror("錯誤", f"找不到Excel檔案A: {xlsx_NewPath_A}")
                return
                
            if not os.path.isfile(xlsx_NewPath_B):
                self._messagebox_showerror("錯誤", f"找不到Excel檔案B: {xlsx_NewPath_B}")
                return
            
            # Construct command
            cmd = f'python3 "{py_NewPath}" "{xlsx_NewPath_A}" "{xlsx_NewPath_B}" "{device_name}" "{Output_NewPath}"'
            
            # Execute command
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            
            # 只顯示 JSON 狀態訊息
            self.cmd_output.delete('1.0', tk.END)
            if result.returncode == 0:
                self.cmd_output.insert('1.0', result.stdout)
                self._messagebox_showinfo("成功", "所有參數檔已成功產生")
            else:
                self.cmd_output.insert('1.0', result.stderr)
                self._messagebox_showerror("錯誤", "執行失敗")
                
        except Exception as e:
            self._messagebox_showerror("錯誤", f"執行時發生錯誤:\n{str(e)}")

def main():
    root = tk.Tk()
    app = XlsxToJsonGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
