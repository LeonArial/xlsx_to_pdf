import os
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext
import comtypes.client
import threading
import queue

class ExcelToPdfConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel 批量转 PDF 工具")
        self.geometry("800x600")

        self.file_list = []
        self.log_queue = queue.Queue()

        # --- UI布局 ---
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件选择和列表区域
        list_frame = ttk.LabelFrame(main_frame, text="待处理文件列表", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.listbox = tk.Listbox(list_frame)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)

        self.select_button = ttk.Button(button_frame, text="选择Excel文件", command=self.select_files)
        self.select_button.pack(side=tk.LEFT, padx=5)

        self.clear_button = ttk.Button(button_frame, text="清空列表", command=self.clear_list)
        self.clear_button.pack(side=tk.LEFT, padx=5)

        self.convert_button = ttk.Button(button_frame, text="开始转换", command=self.start_conversion)
        self.convert_button.pack(side=tk.RIGHT, padx=5)

        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="状态与日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, state='disabled', height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.process_log_queue()

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="请选择一个或多个Excel文件",
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if files:
            for file_path in files:
                if file_path not in self.file_list:
                    self.file_list.append(file_path)
                    self.listbox.insert(tk.END, os.path.basename(file_path))
            self.log(f"已添加 {len(files)} 个文件到列表。")

    def clear_list(self):
        self.file_list.clear()
        self.listbox.delete(0, tk.END)
        self.log("文件列表已清空。")

    def start_conversion(self):
        if not self.file_list:
            self.log("错误：文件列表为空，请先选择文件。")
            return

        self.select_button.config(state='disabled')
        self.clear_button.config(state='disabled')
        self.convert_button.config(state='disabled')
        self.log("=== 开始批量转换 ===")

        # 在新线程中运行转换以避免GUI冻结
        conversion_thread = threading.Thread(
            target=self.batch_convert, 
            args=(self.file_list.copy(),), 
            daemon=True
        )
        conversion_thread.start()

    def batch_convert(self, files_to_convert):
        comtypes.CoInitialize()
        try:
            for i, excel_path in enumerate(files_to_convert):
                pdf_path = os.path.splitext(excel_path)[0] + ".pdf"
                self.log_queue.put(f"({i+1}/{len(files_to_convert)}) 正在处理: {os.path.basename(excel_path)}")
                success = self.convert_excel_to_pdf(excel_path, pdf_path)
                if success:
                    self.log_queue.put(f"  -> 成功转换为: {os.path.basename(pdf_path)}")
                else:
                    self.log_queue.put(f"  -> 转换失败: {os.path.basename(excel_path)}")
        finally:
            self.log_queue.put("=== 批量转换完成 ===")
            comtypes.CoUninitialize()

    def process_log_queue(self):
        try:
            message = self.log_queue.get_nowait()
            self.log(message)
            if message == "=== 批量转换完成 ===":
                self.select_button.config(state='normal')
                self.clear_button.config(state='normal')
                self.convert_button.config(state='normal')
        except queue.Empty:
            pass
        self.after(100, self.process_log_queue)

    def convert_excel_to_pdf(self, excel_path, pdf_path):
        excel = None
        workbook = None
        try:
            excel = comtypes.client.CreateObject("Excel.Application")
            excel.Visible = False
            excel_path_abs = os.path.abspath(excel_path)
            pdf_path_abs = os.path.abspath(pdf_path)

            self.log_queue.put(f"  - 正在打开文件...")
            workbook = excel.Workbooks.Open(excel_path_abs)

            self.log_queue.put("  - 正在设置页面布局...")
            for sheet in workbook.Sheets:
                try:
                    page_setup = sheet.PageSetup
                    page_setup.Orientation = 2  # 横向
                    page_setup.PaperSize = 9    # A4
                    page_setup.Zoom = False
                    page_setup.FitToPagesWide = 1
                    page_setup.FitToPagesTall = False
                except Exception as ps_e:
                    self.log_queue.put(f"  - 警告: 无法为工作表 '{sheet.Name}' 设置页面布局: {ps_e}")

            sheet_names = [sheet.Name for sheet in workbook.Sheets]
            workbook.Worksheets(sheet_names).Select()

            self.log_queue.put(f"  - 正在导出到PDF...")
            workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_path_abs)
            return True
        except Exception as e:
            self.log_queue.put(f"  - 错误: {e}")
            if isinstance(e, comtypes.COMError):
                self.log_queue.put("  - COM错误：请确保已安装Excel且程序有权调用它。")
            return False
        finally:
            if workbook:
                workbook.Close(False)
            if excel:
                excel.Quit()

if __name__ == "__main__":
    app = ExcelToPdfConverterApp()
    app.mainloop()
