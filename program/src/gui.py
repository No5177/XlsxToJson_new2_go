import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading  # 用於創建非阻塞的工作線程
import os
import subprocess
import pandas as pd

class WorkerThread(threading.Thread):
    """工作線程類，用於執行耗時的文件轉換操作"""
    def __init__(self, callback, cmd):
        super().__init__()
        self.callback = callback  # 完成後的回調函數
        self.cmd = cmd  # 要執行的命令
        
    def run(self):
        """執行命令並通過回調返回結果"""
        try:
            result = subprocess.run(self.cmd, shell=True, capture_output=True, text=True)
            self.callback(result)
        except Exception as e:
            self.callback(None, str(e))

class XlsxToJsonGUI:
    """Excel 轉 JSON 的圖形界面類"""
    def __init__(self, root, string_var_class=None, messagebox_showerror=None, messagebox_showinfo=None):
        """初始化 GUI 界面
        
        Args:
            root: tkinter 根窗口
            string_var_class: 字符串變量類（用於測試）
            messagebox_showerror: 錯誤消息框函數（用於測試）
            messagebox_showinfo: 信息消息框函數（用於測試）
        """
        self.root = root
        self.root.title("XLSX to JSON Converter")
        
        # 設置字符串變量類和消息框函數
        self._string_var_class = string_var_class if string_var_class is not None else tk.StringVar
        self._messagebox_showerror = messagebox_showerror if messagebox_showerror is not None else messagebox.showerror
        self._messagebox_showinfo = messagebox_showinfo if messagebox_showinfo is not None else messagebox.showinfo
        
        # 自動設置 xlsx_to_json.py 的路徑
        current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        default_py_path = os.path.join(current_dir, "xlsx_to_json.py")
        
        # 搜尋 Excel 檔案
        root_dir = os.path.dirname(current_dir)  # XlsxToJson 資料夾路徑
        excel_a_path = ""
        excel_b_path = ""
        
        # 搜尋 Excel 檔案
        try:
            for file in os.listdir(root_dir):
                if file.endswith('.xlsx'):
                    # 搜尋 Excel A
                    if file.startswith('POCB-V100A 規格檔 & 保護檔 & 保護回復檔 & 資訊檔一覽表'):
                        excel_a_path = os.path.join(root_dir, file)
                    # 搜尋 Excel B
                    elif file.startswith('POCB-V100A 控制檔一覽表'):
                        excel_b_path = os.path.join(root_dir, file)
        except Exception as e:
            print(f"搜尋 Excel 檔案時發生錯誤: {e}")
        
        # 初始化路徑和設備名稱變量
        self._py_Path = default_py_path if os.path.exists(default_py_path) else ""
        self._xlsx_Path_A = excel_a_path  # Excel A 路徑
        self._xlsx_Path_B = excel_b_path  # Excel B 路徑
        self._device_name = ""
        self._output_path = ""
        
        # 創建 GUI 變量
        self.py_Path = self._string_var_class(value=self._py_Path)
        self.xlsx_Path_A = self._string_var_class(value=self._xlsx_Path_A)
        self.xlsx_Path_B = self._string_var_class(value=self._xlsx_Path_B)
        self.DeviceName = self._string_var_class(value=self._device_name)
        self.Output_Path = self._string_var_class(value=self._output_path)
        
        # 設置變量追踪（當值改變時更新內部變量）
        self.py_Path.trace_add("write", lambda *args: setattr(self, '_py_Path', self.py_Path.get()))
        self.xlsx_Path_A.trace_add("write", lambda *args: setattr(self, '_xlsx_Path_A', self.xlsx_Path_A.get()))
        self.xlsx_Path_B.trace_add("write", lambda *args: setattr(self, '_xlsx_Path_B', self.xlsx_Path_B.get()))
        self.DeviceName.trace_add("write", lambda *args: setattr(self, '_device_name', self.DeviceName.get()))
        self.Output_Path.trace_add("write", lambda *args: setattr(self, '_output_path', self.Output_Path.get()))
        
        # 設置窗口圖標
        try:
            icon_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'icon_image', 'icon_517.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Warning: Could not load icon: {e}")
        
        # 創建主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        # 設置界面元素的固定寬度和間距
        label_width = 25
        entry_width = 50
        button_width = 10
        entry_padx = 5
        button_padx = 5
        row_pady = 5
        
        # 創建各種輸入欄位和按鈕
        # Python 文件選擇
        ttk.Label(main_frame, text="指定python檔", width=label_width).grid(
            row=0, column=0, sticky=tk.W, pady=row_pady)
        ttk.Entry(main_frame, textvariable=self.py_Path, width=entry_width).grid(
            row=0, column=1, padx=entry_padx, sticky=tk.EW)
        ttk.Button(main_frame, text="瀏覽", command=self.select_py_file, width=button_width).grid(
            row=0, column=2, padx=button_padx)
        
        # Excel file A selection
        ttk.Label(main_frame, text="指定Excel檔案(規格檔 & 保護檔 & 保護回復檔 & 資訊檔一覽表)", width=label_width).grid(
            row=1, column=0, sticky=tk.W, pady=row_pady)
        ttk.Entry(main_frame, textvariable=self.xlsx_Path_A, width=entry_width).grid(
            row=1, column=1, padx=entry_padx, sticky=tk.EW)
        ttk.Button(main_frame, text="瀏覽", command=self.select_xlsx_file_A, width=button_width).grid(
            row=1, column=2, padx=button_padx)
        
        # Excel file B selection
        ttk.Label(main_frame, text="指定Excel檔案(控制檔一覽表)", width=label_width).grid(
            row=2, column=0, sticky=tk.W, pady=row_pady)
        ttk.Entry(main_frame, textvariable=self.xlsx_Path_B, width=entry_width).grid(
            row=2, column=1, padx=entry_padx, sticky=tk.EW)
        ttk.Button(main_frame, text="瀏覽", command=self.select_xlsx_file_B, width=button_width).grid(
            row=2, column=2, padx=button_padx)
        
        # Device name input
        ttk.Label(main_frame, text="指定設備型號", width=label_width).grid(
            row=3, column=0, sticky=tk.W, pady=row_pady)
        self.device_combobox = ttk.Combobox(main_frame, textvariable=self.DeviceName, width=entry_width)
        self.device_combobox.grid(row=3, column=1, padx=entry_padx, sticky=tk.EW)
        ttk.Button(main_frame, text="讀取型號", command=self.load_device_names, width=button_width).grid(
            row=3, column=2, padx=button_padx)
        
        # Output path selection
        ttk.Label(main_frame, text="指定產出檔案路徑", width=label_width).grid(
            row=4, column=0, sticky=tk.W, pady=row_pady)
        ttk.Entry(main_frame, textvariable=self.Output_Path, width=entry_width).grid(
            row=4, column=1, padx=entry_padx, sticky=tk.EW)
        ttk.Button(main_frame, text="瀏覽", command=self.select_output_path, width=button_width).grid(
            row=4, column=2, padx=button_padx)
        
        # Command output display
        ttk.Label(main_frame, text="命令執行狀態", width=label_width).grid(row=5, column=0, sticky=tk.W, pady=5)
        self.cmd_output = tk.Text(main_frame, height=10, width=70)
        self.cmd_output.grid(row=5, column=1, columnspan=2, padx=entry_padx, pady=5)
        
        # 修改进度条位置和大小
        self.progress_bar = ttk.Progressbar(main_frame, mode='determinate')
        self.progress_bar.grid(row=6, column=0, columnspan=3, sticky="ew", padx=(0, 5))  # 調整對齊方式
        
        # Generate button (移到进度条下方)
        ttk.Button(main_frame, text="所有參數檔生成", command=self.generate_fw_spec).grid(row=7, column=1, pady=20)
        
    def select_py_file(self):
        """選擇 Python 文件的對話框"""
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
    
    def load_device_names(self):
        """讀取 Excel A 中的設備型號列表"""
        try:
            # 檢查是否已選擇 Excel A
            if not self.xlsx_Path_A.get():
                messagebox.showerror("錯誤", "請先選擇 Excel A 檔案")
                return
            
            # 讀取 Excel 檔案
            df = pd.read_excel(
                self.xlsx_Path_A.get(),
                sheet_name="LTRp 規格檔一覽表",
                engine='openpyxl',
                header=None
            )
            
            # 獲取第4行（索引3）的設備型號
            device_list = df.iloc[3, 4:].dropna().unique()  # 從第5列開始（跳過前4列）
            
            # 更新下拉選單
            self.device_combobox['values'] = list(device_list)
            
            # 如果找到型號，顯示成功訊息
            if len(device_list) > 0:
                messagebox.showinfo("成功", f"已讀取 {len(device_list)} 個設備型號")
            else:
                messagebox.showwarning("警告", "未找到任何設備型號")
                
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取設備型號失敗：{str(e)}")
    
    def validate_inputs(self):
        """驗證輸入"""
        # ... 其他驗證保持不變 ...
        
        # 驗證設備型號
        if not self.DeviceName.get():
            messagebox.showerror("錯誤", "請選擇設備型號")
            return False
        
        # ... 其他驗證保持不變 ...
        return True
    
    def generate_fw_spec(self):
        """生成固件規格文件的主要函數"""
        try:
            # 獲取並清理輸入值
            py_path = str(self.py_Path.get()).strip()
            xlsx_path_a = str(self.xlsx_Path_A.get()).strip()
            xlsx_path_b = str(self.xlsx_Path_B.get()).strip()
            device_name = str(self.DeviceName.get()).strip()
            output_path = str(self.Output_Path.get()).strip()
            
            # 驗證輸入
            if not all([py_path, xlsx_path_a, xlsx_path_b, device_name, output_path]):
                self._messagebox_showerror("錯誤", "請填寫所有必要欄位")
                return
            
            # 重置進度條
            self.progress_bar['value'] = 0
            
            # 使用絕對路徑
            current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            
            # Convert paths to absolute paths
            py_NewPath = os.path.abspath(py_path)
            xlsx_NewPath_A = os.path.abspath(xlsx_path_a)
            xlsx_NewPath_B = os.path.abspath(xlsx_path_b)
            Output_NewPath = os.path.abspath(output_path)
            
            # 構建命令時使用 python 而不是 python3
            cmd = f'python "{py_NewPath}" "{xlsx_NewPath_A}" "{xlsx_NewPath_B}" "{device_name}" "{Output_NewPath}"'
            
            # 更新進度條到 50%
            self.progress_bar['value'] = 50
            self.root.update_idletasks()
            
            # 定義完成處理的回調函數
            def process_complete(result, error=None):
                """處理完成後的回調函數"""
                if error:
                    print(f"Error: {error}")  # 添加命令行輸出
                    self.cmd_output.delete('1.0', tk.END)
                    self.cmd_output.insert('1.0', f"錯誤: {error}")
                    self._messagebox_showerror("錯誤", f"執行時發生錯誤:\n{error}")
                    self.progress_bar['value'] = 0
                elif result.returncode == 0:
                    print("Success")  # 添加命令行輸出
                    self.cmd_output.delete('1.0', tk.END)
                    self.cmd_output.insert('1.0', result.stdout)
                    self._messagebox_showinfo("成功", "所有參數檔已成功產生")
                    self.progress_bar['value'] = 100
                else:
                    print(f"Error output: {result.stderr}")  # 添加命令行輸出
                    error_message = result.stderr if result.stderr else result.stdout
                    self.cmd_output.delete('1.0', tk.END)
                    self.cmd_output.insert('1.0', error_message)
                    self._messagebox_showerror("錯誤", f"執行失敗:\n{error_message}")
                    self.progress_bar['value'] = 0
            
            # 創建並啟動工作線程
            worker = WorkerThread(process_complete, cmd)
            worker.start()
                
        except Exception as e:
            error_msg = str(e)
            print(f"Error: {error_msg}")
            self.cmd_output.delete('1.0', tk.END)
            self.cmd_output.insert('1.0', f"錯誤: {error_msg}")
            self._messagebox_showerror("錯誤", f"執行時發生錯誤:\n{error_msg}")
            self.progress_bar['value'] = 0

def main():
    """主函數，創建並運行 GUI 應用"""
    root = tk.Tk()
    app = XlsxToJsonGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
