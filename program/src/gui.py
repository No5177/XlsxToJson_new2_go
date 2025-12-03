import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading  # 用於創建非阻塞的工作線程
import os
import sys
import io
import pandas as pd

from pathlib import Path

# 确保能够导入上级目录的模块
parent_dir = str(Path(__file__).resolve().parent.parent)
if parent_dir not in sys.path:
    sys.path.append(parent_dir)

# 直接导入 xlsx_to_json 模块
import xlsx_to_json

# 資源路徑處理函數，用於打包後正確讀取資源檔案
def resource_path(relative_path):
    """獲取資源的絕對路徑，兼容開發環境和打包後的執行環境"""
    try:
        # PyInstaller 和 Nuitka 打包後的臨時目錄
        base_path = getattr(sys, '_MEIPASS', None)
        if base_path is None:
            # 嘗試 Nuitka 特有的環境變量
            base_path = os.environ.get('NUITKA_ONEFILE_PARENT', None)
            
        if base_path is None:
            # 開發環境或打包後的執行檔所在目錄
            base_path = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
    except Exception:
        # 如果上述方法都失敗，使用當前目錄
        base_path = os.path.abspath(".")
        
    return os.path.join(base_path, relative_path)

class WorkerThread(threading.Thread):
    """工作線程類，用於執行耗時的文件轉換操作"""
    def __init__(self, callback, func, *args, **kwargs):
        super().__init__()
        self.callback = callback  # 完成後的回調函數
        self.func = func  # 要執行的函数
        self.args = args
        self.kwargs = kwargs
        
    def run(self):
        """執行函數並通過回調返回結果"""
        try:
            # 捕获标准输出和错误
            original_stdout = sys.stdout
            original_stderr = sys.stderr
            stdout_capture = io.StringIO()
            stderr_capture = io.StringIO()
            
            sys.stdout = stdout_capture
            sys.stderr = stderr_capture
            
            try:
                self.func(*self.args, **self.kwargs)
                success = True
            except Exception as e:
                success = False
                error = str(e)
            finally:
                # 恢复标准输出和错误
                sys.stdout = original_stdout
                sys.stderr = original_stderr
            
            if success:
                self.callback(True, stdout_capture.getvalue(), None)
            else:
                self.callback(False, stdout_capture.getvalue(), stderr_capture.getvalue() or error)
        except Exception as e:
            self.callback(False, "", str(e))

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
        
        # 搜尋 Excel 檔案
        # 獲取執行檔所在目錄
        if getattr(sys, 'frozen', False):
            # 打包後的環境
            exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        else:
            # 開發環境
            exe_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        
        excel_a_path = ""
        excel_b_path = ""
        
        # 搜尋 Excel 檔案
        try:
            # 首先在執行檔所在目錄搜尋
            for file in os.listdir(exe_dir):
                if file.endswith('.xlsx'):
                    # 搜尋 Excel A
                    if file.startswith('POCB-V100A 規格檔 & 保護檔 & 保護回復檔 & 資訊檔一覽表'):
                        excel_a_path = os.path.join(exe_dir, file)
                    # 搜尋 Excel B
                    elif file.startswith('POCB-V100A 控制檔一覽表'):
                        excel_b_path = os.path.join(exe_dir, file)
            
            # 如果在執行檔所在目錄沒找到，嘗試在其他常見位置搜尋
            if not excel_a_path or not excel_b_path:
                other_dirs = [
                    os.path.join(exe_dir, "excel"),
                    os.path.join(exe_dir, "data"),
                    os.path.expanduser("~/Documents"),
                    os.path.abspath(".")
                ]
                
                for dir_path in other_dirs:
                    if os.path.exists(dir_path):
                        for file in os.listdir(dir_path):
                            if file.endswith('.xlsx'):
                                if file.startswith('POCB-V100A 規格檔 & 保護檔 & 保護回復檔 & 資訊檔一覽表') and not excel_a_path:
                                    excel_a_path = os.path.join(dir_path, file)
                                elif file.startswith('POCB-V100A 控制檔一覽表') and not excel_b_path:
                                    excel_b_path = os.path.join(dir_path, file)
        except Exception as e:
            print(f"搜尋 Excel 檔案時發生錯誤: {e}")
        
        # 初始化路徑和設備名稱變量
        self._xlsx_Path_A = excel_a_path  # Excel A 路徑
        self._xlsx_Path_B = excel_b_path  # Excel B 路徑
        self._device_name = ""
        self._output_path = ""
        
        # 創建 GUI 變量
        self.xlsx_Path_A = self._string_var_class(value=self._xlsx_Path_A)
        self.xlsx_Path_B = self._string_var_class(value=self._xlsx_Path_B)
        self.DeviceName = self._string_var_class(value=self._device_name)
        self.Output_Path = self._string_var_class(value=self._output_path)
        
        # 設置變量追踪（當值改變時更新內部變量）
        self.xlsx_Path_A.trace_add("write", lambda *args: setattr(self, '_xlsx_Path_A', self.xlsx_Path_A.get()))
        self.xlsx_Path_B.trace_add("write", lambda *args: setattr(self, '_xlsx_Path_B', self.xlsx_Path_B.get()))
        self.DeviceName.trace_add("write", lambda *args: setattr(self, '_device_name', self.DeviceName.get()))
        self.Output_Path.trace_add("write", lambda *args: setattr(self, '_output_path', self.Output_Path.get()))
        
        # 設置窗口圖標
        try:
            # 先嘗試使用 resource_path 函數查找 icon.ico
            icon_path = resource_path("icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
            else:
                # 後備方案：嘗試在 icon_image 目錄中查找
                icon_path = resource_path(os.path.join("icon_image", "icon_517.ico"))
                if os.path.exists(icon_path):
                    self.root.iconbitmap(icon_path)
                else:
                    # 再次後備：使用舊的路徑方式查找
                    icon_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'icon_image', 'icon_517.ico')
                    if os.path.exists(icon_path):
                        self.root.iconbitmap(icon_path)
                    else:
                        print("Warning: Could not find any icon file.")
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
        # Excel file A selection
        ttk.Label(main_frame, text="指定Excel檔案(規格檔 & 保護檔 & 保護回復檔 & 資訊檔一覽表)", width=label_width).grid(
            row=0, column=0, sticky=tk.W, pady=row_pady)
        ttk.Entry(main_frame, textvariable=self.xlsx_Path_A, width=entry_width).grid(
            row=0, column=1, padx=entry_padx, sticky=tk.EW)
        ttk.Button(main_frame, text="瀏覽", command=self.select_xlsx_file_A, width=button_width).grid(
            row=0, column=2, padx=button_padx)
        
        # Excel file B selection
        ttk.Label(main_frame, text="指定Excel檔案(控制檔一覽表)", width=label_width).grid(
            row=1, column=0, sticky=tk.W, pady=row_pady)
        ttk.Entry(main_frame, textvariable=self.xlsx_Path_B, width=entry_width).grid(
            row=1, column=1, padx=entry_padx, sticky=tk.EW)
        ttk.Button(main_frame, text="瀏覽", command=self.select_xlsx_file_B, width=button_width).grid(
            row=1, column=2, padx=button_padx)
        
        # Device name input
        ttk.Label(main_frame, text="指定設備型號", width=label_width).grid(
            row=2, column=0, sticky=tk.W, pady=row_pady)
        self.device_combobox = ttk.Combobox(main_frame, textvariable=self.DeviceName, width=entry_width)
        self.device_combobox.grid(row=2, column=1, padx=entry_padx, sticky=tk.EW)
        ttk.Button(main_frame, text="讀取型號", command=self.load_device_names, width=button_width).grid(
            row=2, column=2, padx=button_padx)
        
        # Output path selection
        ttk.Label(main_frame, text="指定產出檔案路徑", width=label_width).grid(
            row=3, column=0, sticky=tk.W, pady=row_pady)
        ttk.Entry(main_frame, textvariable=self.Output_Path, width=entry_width).grid(
            row=3, column=1, padx=entry_padx, sticky=tk.EW)
        ttk.Button(main_frame, text="瀏覽", command=self.select_output_path, width=button_width).grid(
            row=3, column=2, padx=button_padx)
        
        # Command output display
        ttk.Label(main_frame, text="執行狀態", width=label_width).grid(row=4, column=0, sticky=tk.W, pady=5)
        self.cmd_output = tk.Text(main_frame, height=10, width=70)
        self.cmd_output.grid(row=4, column=1, columnspan=2, padx=entry_padx, pady=5)
        
        # 修改进度条位置和大小
        self.progress_bar = ttk.Progressbar(main_frame, mode='determinate')
        self.progress_bar.grid(row=5, column=0, columnspan=3, sticky="ew", padx=(0, 5))  # 調整對齊方式
        
        # Generate button (移到进度条下方)
        ttk.Button(main_frame, text="所有參數檔生成", command=self.generate_fw_spec).grid(row=6, column=1, pady=20)
    
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
        # 驗證Excel檔案A
        if not self.xlsx_Path_A.get():
            self._messagebox_showerror("錯誤", "請選擇Excel檔案A (規格檔&保護檔&保護回復檔&資訊檔)")
            return False
            
        # 驗證Excel檔案B
        if not self.xlsx_Path_B.get():
            self._messagebox_showerror("錯誤", "請選擇Excel檔案B (控制檔)")
            return False
        
        # 驗證設備型號
        if not self.DeviceName.get():
            self._messagebox_showerror("錯誤", "請選擇設備型號")
            return False
        
        # 驗證輸出路徑
        if not self.Output_Path.get():
            self._messagebox_showerror("錯誤", "請選擇輸出目錄")
            return False
            
        return True
    
    def generate_fw_spec(self):
        """生成固件規格文件的主要函數"""
        try:
            # 獲取並清理輸入值
            xlsx_path_a = str(self.xlsx_Path_A.get()).strip()
            xlsx_path_b = str(self.xlsx_Path_B.get()).strip()
            device_name = str(self.DeviceName.get()).strip()
            output_path = str(self.Output_Path.get()).strip()
            
            # 驗證輸入
            if not self.validate_inputs():
                return
            
            # 重置進度條和輸出文字
            self.progress_bar['value'] = 0
            self.cmd_output.delete('1.0', tk.END)
            self.cmd_output.insert('1.0', "開始處理中...\n")
            
            # 轉換為絕對路徑
            xlsx_path_a = os.path.abspath(xlsx_path_a)
            xlsx_path_b = os.path.abspath(xlsx_path_b)
            output_path = os.path.abspath(output_path)
            
            # 更新進度條到 25%
            self.progress_bar['value'] = 25
            self.root.update_idletasks()
            
            # 定義完成處理的回調函數
            def process_complete(success, stdout, stderr):
                """處理完成後的回調函數"""
                if not success:
                    print(f"Error: {stderr}")  # 添加命令行輸出
                    self.cmd_output.delete('1.0', tk.END)
                    self.cmd_output.insert('1.0', f"錯誤: {stderr}")
                    self._messagebox_showerror("錯誤", f"執行時發生錯誤:\n{stderr}")
                    self.progress_bar['value'] = 0
                else:
                    print("Success")  # 添加命令行輸出
                    self.cmd_output.delete('1.0', tk.END)
                    self.cmd_output.insert('1.0', stdout)
                    self._messagebox_showinfo("成功", "所有參數檔已成功產生")
                    self.progress_bar['value'] = 100
            
            # 創建並啟動工作線程，直接呼叫 xlsx_to_json 模組的功能
            worker = WorkerThread(
                process_complete, 
                xlsx_to_json.convert_xlsx_to_json,
                xlsx_path_a,
                xlsx_path_b,
                device_name,
                output_path
            )
            worker.start()
            
            # 更新進度條到 50%
            self.progress_bar['value'] = 50
            self.root.update_idletasks()
                
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
