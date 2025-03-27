import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk  # 添加ttk导入
import pandas as pd
import json
import os

class ExcelToJsonConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 转 JSON 工具【MDC小助手】")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        self.excel_path = ""
        self.json_path = ""
        
        self.create_widgets()
        
    def create_widgets(self):
        # 顶部框架 - 文件选择
        top_frame = tk.Frame(self.root, pady=10)
        top_frame.pack(fill=tk.X)
        
        # Excel 文件选择
        tk.Label(top_frame, text="Excel 文件:").grid(row=0, column=0, padx=5, sticky=tk.W)
        self.excel_entry = tk.Entry(top_frame, width=50)
        self.excel_entry.grid(row=0, column=1, padx=5)
        tk.Button(top_frame, text="浏览...", command=self.browse_excel).grid(row=0, column=2, padx=5)
        
        # JSON 文件选择
        tk.Label(top_frame, text="JSON 输出:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.json_entry = tk.Entry(top_frame, width=50)
        self.json_entry.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(top_frame, text="浏览...", command=self.browse_json).grid(row=1, column=2, padx=5, pady=5)
        
        # 转换按钮
        convert_button = tk.Button(top_frame, text="转换", command=self.convert, bg="#4CAF50", fg="white", padx=20)
        convert_button.grid(row=2, column=1, pady=10)
        
        # 中间框架 - 选项
        mid_frame = tk.LabelFrame(self.root, text="选项", padx=10, pady=10)
        mid_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 选项 - 工作表选择
        tk.Label(mid_frame, text="工作表名称:").grid(row=0, column=0, sticky=tk.W, padx=5)
        # 将Entry替换为Combobox
        self.sheet_combobox = ttk.Combobox(mid_frame, width=18)
        self.sheet_combobox.grid(row=0, column=1, sticky=tk.W, padx=5)
        # 不再需要 self.sheet_entry.insert(0, "Sheet1")
        
        # 选项 - 是否使用第一行作为键
        self.use_header_var = tk.BooleanVar(value=True)
        tk.Checkbutton(mid_frame, text="使用第一行作为键", variable=self.use_header_var).grid(row=0, column=2, padx=20, sticky=tk.W)
        
        # 选项 - 输出格式
        tk.Label(mid_frame, text="输出格式:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.format_var = tk.StringVar(value="records")
        tk.Radiobutton(mid_frame, text="记录列表", variable=self.format_var, value="records").grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        tk.Radiobutton(mid_frame, text="列表字典", variable=self.format_var, value="list").grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
        
        # 底部框架 - 结果显示
        bottom_frame = tk.LabelFrame(self.root, text="转换结果", padx=10, pady=10)
        bottom_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 结果文本框
        self.result_text = scrolledtext.ScrolledText(bottom_frame, wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        # 添加MDC水印
        watermark_label = tk.Label(bottom_frame, text="MDC", fg="lightgray", font=("Arial", 36, "bold"))
        watermark_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("准备就绪")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def browse_excel(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.excel_path = file_path
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, file_path)
            
            # 自动设置JSON输出路径
            base_name = os.path.splitext(file_path)[0]
            self.json_path = f"{base_name}.json"
            self.json_entry.delete(0, tk.END)
            self.json_entry.insert(0, self.json_path)
            
            # 获取工作表列表并更新下拉菜单
            try:
                # 使用ExcelFile获取所有工作表名称
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
                sheet_names = excel_file.sheet_names
                
                # 更新工作表下拉菜单
                self.sheet_combobox['values'] = sheet_names
                if sheet_names:
                    self.sheet_combobox.current(0)  # 选择第一个工作表
            except Exception as e:
                messagebox.showerror("错误", f"读取工作表列表时出错:\n{str(e)}")
    
    def browse_json(self):
        file_path = filedialog.asksaveasfilename(
            title="保存JSON文件",
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if file_path:
            self.json_path = file_path
            self.json_entry.delete(0, tk.END)
            self.json_entry.insert(0, file_path)
    
    def convert(self):
        try:
            excel_path = self.excel_entry.get()
            json_path = self.json_entry.get()
            sheet_name = self.sheet_combobox.get()  # 从下拉菜单获取工作表名称
            
            if not excel_path:
                messagebox.showerror("错误", "请选择Excel文件")
                return
            
            if not json_path:
                messagebox.showerror("错误", "请选择JSON输出位置")
                return
            
            if not sheet_name:
                messagebox.showerror("错误", "请选择工作表")
                return
            
            self.status_var.set("正在转换...")
            self.root.update()
            
            # 读取Excel文件
            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine='openpyxl')
            
            # 转换为JSON
            if self.format_var.get() == "records":
                # 记录列表格式 [{col1:val1, col2:val2}, {col1:val3, col2:val4}]
                json_data = df.to_json(orient="records", force_ascii=False)
            else:
                # 列表字典格式 {col1:[val1,val3], col2:[val2,val4]}
                json_data = df.to_json(orient="columns", force_ascii=False)
            
            # 格式化JSON以便显示
            parsed_json = json.loads(json_data)
            formatted_json = json.dumps(parsed_json, indent=4, ensure_ascii=False)
            
            # 保存到文件
            with open(json_path, 'w', encoding='utf-8') as f:
                f.write(formatted_json)
            
            # 显示结果
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, formatted_json)
            
            self.status_var.set(f"转换完成! 已保存到 {json_path}")
            messagebox.showinfo("成功", f"Excel已成功转换为JSON并保存到:\n{json_path}")
            
        except Exception as e:
            self.status_var.set(f"转换失败: {str(e)}")
            messagebox.showerror("错误", f"转换过程中发生错误:\n{str(e)}")
            
            # 显示详细错误信息
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"错误详情:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToJsonConverter(root)
    root.mainloop()