# -*- coding: utf-8 -*-
"""
GUI界面模块
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
from datetime import datetime

from parser_core import AIDA64Parser
from templates import TemplateManager


class AIDA64ParserApp:
    """AIDA64解析器应用程序"""
    
    def __init__(self):
        self.parser = AIDA64Parser()
        self.template_manager = TemplateManager()
        self.selected_files = []
        self.parsed_data = {}
        
        # 创建主窗口
        self.root = tk.Tk()
        self.root.title("AIDA64报告智能解析工具 v1.0")
        self.root.geometry("1000x700")
        self.root.minsize(900, 600)
        
        # 设置图标（如果有）
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
        
        # 创建界面
        self.create_widgets()
        
        # 设置样式
        self.setup_styles()
    
    def setup_styles(self):
        """设置界面样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置颜色
        self.root.configure(bg='#f0f0f0')
        
        # 配置按钮样式
        style.configure('TButton', padding=6)
        style.configure('Title.TLabel', font=('Arial', 14, 'bold'))
        style.configure('Subtitle.TLabel', font=('Arial', 10, 'bold'))
    
    def create_widgets(self):
        """创建界面组件"""
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="AIDA64报告智能解析工具", 
                               style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))
        
        # 工具栏
        toolbar = ttk.Frame(main_frame)
        toolbar.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 工具栏按钮
        ttk.Button(toolbar, text="选择文件", command=self.select_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="选择文件夹", command=self.select_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="清空列表", command=self.clear_file_list).pack(side=tk.LEFT, padx=2)
        
        tk.Frame(toolbar, width=20).pack(side=tk.LEFT)  # 间隔
        
        ttk.Button(toolbar, text="开始解析", command=self.start_parsing, 
                  style='Accent.TButton').pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="导出Excel", command=self.export_excel).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="导出所有格式", command=self.export_all).pack(side=tk.LEFT, padx=2)
        
        # 文件列表和配置区域
        left_frame = ttk.LabelFrame(main_frame, text="文件列表", padding="5")
        left_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        right_frame = ttk.LabelFrame(main_frame, text="解析配置", padding="5")
        right_frame.grid(row=2, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        
        # 配置网格权重
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(1, weight=1)
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(1, weight=1)
        
        # 文件列表
        file_list_frame = ttk.Frame(left_frame)
        file_list_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        ttk.Label(file_list_frame, text=f"已选择 {len(self.selected_files)} 个文件").pack(side=tk.LEFT)
        
        # 文件列表框
        self.file_listbox = tk.Listbox(left_frame, selectmode=tk.EXTENDED, 
                                      bg='white', relief=tk.SUNKEN, borderwidth=1)
        self.file_listbox.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 0))
        
        # 文件列表框滚动条
        file_scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, 
                                      command=self.file_listbox.yview)
        file_scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S), pady=(5, 0))
        self.file_listbox.config(yscrollcommand=file_scrollbar.set)
        
        # 模板选择
        template_frame = ttk.Frame(right_frame)
        template_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(template_frame, text="选择模板:").pack(side=tk.LEFT)
        
        self.template_var = tk.StringVar(value='standard')
        template_combo = ttk.Combobox(template_frame, textvariable=self.template_var,
                                     values=self.template_manager.get_available_templates(),
                                     state='readonly', width=20)
        template_combo.pack(side=tk.LEFT, padx=5)
        
        # 自定义项目选项
        options_frame = ttk.Frame(right_frame)
        options_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.include_software_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="包含已安装程序", 
                       variable=self.include_software_var).pack(side=tk.LEFT, padx=5)
        
        self.custom_items_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="自定义提取项目", 
                       variable=self.custom_items_var,
                       command=self.toggle_custom_items).pack(side=tk.LEFT, padx=5)
        
        # 自定义项目输入
        self.custom_items_frame = ttk.Frame(right_frame)
        self.custom_items_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(self.custom_items_frame, text="输入项目（每行一个）:").pack(anchor=tk.W)
        
        self.custom_items_text = scrolledtext.ScrolledText(self.custom_items_frame, 
                                                          height=8, width=40)
        self.custom_items_text.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # 示例项目
        example_items = "计算机类型\n操作系统\n处理器名称\n主板名称\n系统内存\n显示适配器"
        self.custom_items_text.insert('1.0', example_items)
        
        self.custom_items_frame.grid_remove()  # 默认隐藏
        
        # 输出格式选择
        format_frame = ttk.Frame(right_frame)
        format_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(format_frame, text="输出格式:").pack(side=tk.LEFT)
        
        self.format_var = tk.StringVar(value='excel')
        ttk.Radiobutton(format_frame, text="Excel", variable=self.format_var, 
                       value='excel').pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(format_frame, text="CSV", variable=self.format_var, 
                       value='csv').pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(format_frame, text="JSON", variable=self.format_var, 
                       value='json').pack(side=tk.LEFT, padx=5)
        
        # 输出文件名
        output_frame = ttk.Frame(right_frame)
        output_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(output_frame, text="输出文件名:").pack(side=tk.LEFT)
        
        self.output_name_var = tk.StringVar(value=f"aida64_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        output_entry = ttk.Entry(output_frame, textvariable=self.output_name_var, width=30)
        output_entry.pack(side=tk.LEFT, padx=5)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="操作日志", padding="5")
        log_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, state='disabled')
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
    
    def select_files(self):
        """选择文件"""
        files = filedialog.askopenfilenames(
            title="选择AIDA64报告文件",
            filetypes=[
                ("文本文件", "*.txt"),
                ("所有文件", "*.*")
            ]
        )
        
        if files:
            self.selected_files.extend(files)
            self.update_file_list()
            self.log(f"已添加 {len(files)} 个文件")
    
    def select_folder(self):
        """选择文件夹"""
        folder = filedialog.askdirectory(title="选择包含AIDA64报告的文件夹")
        
        if folder:
            # 查找所有txt文件
            txt_files = []
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith('.txt'):
                        txt_files.append(os.path.join(root, file))
            
            if txt_files:
                self.selected_files.extend(txt_files)
                self.update_file_list()
                self.log(f"从文件夹添加 {len(txt_files)} 个文件")
            else:
                messagebox.showwarning("警告", "选择的文件夹中没有找到.txt文件")
    
    def clear_file_list(self):
        """清空文件列表"""
        self.selected_files = []
        self.file_listbox.delete(0, tk.END)
        self.log("已清空文件列表")
    
    def update_file_list(self):
        """更新文件列表显示"""
        self.file_listbox.delete(0, tk.END)
        
        for file_path in self.selected_files:
            filename = os.path.basename(file_path)
            self.file_listbox.insert(tk.END, filename)
        
        # 更新文件计数
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.LabelFrame):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Frame):
                        for grandchild in child.winfo_children():
                            if isinstance(grandchild, ttk.Label):
                                if "已选择" in grandchild.cget("text"):
                                    grandchild.config(text=f"已选择 {len(self.selected_files)} 个文件")
    
    def toggle_custom_items(self):
        """切换自定义项目显示"""
        if self.custom_items_var.get():
            self.custom_items_frame.grid()
        else:
            self.custom_items_frame.grid_remove()
    
    def get_selected_items(self):
        """获取选中的项目列表"""
        template_name = self.template_var.get()
        
        if self.custom_items_var.get():
            # 使用自定义项目
            custom_text = self.custom_items_text.get('1.0', tk.END).strip()
            items = [line.strip() for line in custom_text.split('\n') if line.strip()]
            return items
        else:
            # 使用模板项目
            return self.template_manager.get_template_items(template_name)
    
    def start_parsing(self):
        """开始解析"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先选择要解析的文件")
            return
        
        # 获取选中的项目
        selected_items = self.get_selected_items()
        
        # 更新状态
        self.status_var.set("正在解析...")
        self.log(f"开始解析 {len(self.selected_files)} 个文件")
        self.log(f"使用模板: {self.template_var.get()}")
        
        # 在新线程中执行解析
        threading.Thread(target=self._parse_in_thread, 
                        args=(selected_items,), daemon=True).start()
    
    def _parse_in_thread(self, selected_items):
        """在新线程中执行解析"""
        try:
            # 解析文件
            self.parsed_data = self.parser.parse_multiple_files(
                self.selected_files, selected_items
            )
            
            # 更新UI
            self.root.after(0, self._on_parsing_complete)
            
        except Exception as e:
            self.root.after(0, lambda: self._on_parsing_error(str(e)))
    
    def _on_parsing_complete(self):
        """解析完成"""
        total_items = sum(len(data) for data in self.parsed_data.values())
        self.status_var.set(f"解析完成，共提取 {total_items} 条数据")
        self.log(f"解析完成，共 {len(self.parsed_data)} 个文件，{total_items} 条数据")
        
        # 显示预览
        self.show_preview()
    
    def _on_parsing_error(self, error_msg):
        """解析出错"""
        self.status_var.set("解析出错")
        self.log(f"解析错误: {error_msg}")
        messagebox.showerror("错误", f"解析过程中出现错误:\n{error_msg}")
    
    def show_preview(self):
        """显示预览"""
        if not self.parsed_data:
            return
        
        # 创建预览窗口
        preview_window = tk.Toplevel(self.root)
        preview_window.title("解析结果预览")
        preview_window.geometry("800x600")
        
        # 创建笔记本控件
        notebook = ttk.Notebook(preview_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 为每个文件添加标签页
        for filename, data in self.parsed_data.items():
            frame = ttk.Frame(notebook)
            notebook.add(frame, text=filename)
            
            # 创建树状视图
            columns = ('项目', '值')
            tree = ttk.Treeview(frame, columns=columns, show='headings')
            
            # 设置列标题
            tree.heading('项目', text='项目')
            tree.heading('值', text='值')
            
            # 设置列宽
            tree.column('项目', width=300)
            tree.column('值', width=400)
            
            # 添加滚动条
            scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            
            # 布局
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # 添加数据
            for item in data:
                tree.insert('', tk.END, values=(item['项目'], item['值']))
    
    def export_excel(self):
        """导出为Excel"""
        if not self.parsed_data:
            messagebox.showwarning("警告", "请先解析文件")
            return
        
        # 选择保存位置
        default_name = self.output_name_var.get() + '.xlsx'
        file_path = filedialog.asksaveasfilename(
            title="保存Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialfile=default_name
        )
        
        if file_path:
            try:
                self.status_var.set("正在导出Excel...")
                self.log(f"导出Excel到: {file_path}")
                
                # 在新线程中执行导出
                threading.Thread(target=self._export_excel_thread, 
                               args=(file_path,), daemon=True).start()
                
            except Exception as e:
                messagebox.showerror("错误", f"导出失败:\n{str(e)}")
                self.log(f"导出失败: {str(e)}")
    
    def _export_excel_thread(self, file_path):
        """在新线程中导出Excel"""
        try:
            output_path = self.parser.export_to_excel(self.parsed_data, file_path)
            self.root.after(0, lambda: self._on_export_complete(output_path, 'Excel'))
        except Exception as e:
            self.root.after(0, lambda: self._on_export_error(str(e)))
    
    def export_all(self):
        """导出所有格式"""
        if not self.parsed_data:
            messagebox.showwarning("警告", "请先解析文件")
            return
        
        # 选择保存文件夹
        folder = filedialog.askdirectory(title="选择保存文件夹")
        if not folder:
            return
        
        try:
            base_name = self.output_name_var.get()
            
            # 导出Excel
            excel_path = os.path.join(folder, f"{base_name}.xlsx")
            self.parser.export_to_excel(self.parsed_data, excel_path)
            
            # 导出CSV（为每个文件单独导出）
            for filename, data in self.parsed_data.items():
                import pandas as pd
                csv_filename = f"{base_name}_{os.path.splitext(filename)[0]}.csv"
                csv_path = os.path.join(folder, csv_filename)
                
                df = pd.DataFrame(data)
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            
            # 导出JSON
            json_path = os.path.join(folder, f"{base_name}.json")
            import json
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(self.parsed_data, f, ensure_ascii=False, indent=2)
            
            self.log(f"已导出所有格式到: {folder}")
            messagebox.showinfo("成功", f"已成功导出所有格式到:\n{folder}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败:\n{str(e)}")
            self.log(f"导出失败: {str(e)}")
    
    def _on_export_complete(self, file_path, format_name):
        """导出完成"""
        self.status_var.set(f"{format_name}导出完成")
        self.log(f"已成功导出到: {file_path}")
        
        # 询问是否打开文件
        if messagebox.askyesno("成功", f"已成功导出{format_name}文件:\n{file_path}\n\n是否要打开文件？"):
            os.startfile(file_path)
    
    def _on_export_error(self, error_msg):
        """导出出错"""
        self.status_var.set("导出出错")
        self.log(f"导出错误: {error_msg}")
        messagebox.showerror("错误", f"导出过程中出现错误:\n{error_msg}")
    
    def log(self, message):
        """添加日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
    
    def run(self):
        """运行应用程序"""
        # 居中显示窗口
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
        self.root.mainloop()


if __name__ == "__main__":
    app = AIDA64ParserApp()
    app.run()
