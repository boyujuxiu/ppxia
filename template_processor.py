import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os
import shutil
import json
from pathlib import Path
from datetime import datetime

class TemplateProcessor:
    def __init__(self):
        # 添加异常捕获
        try:
            self.window = tk.Tk()
        except Exception as e:
            print(f"初始化失败: {str(e)}")
            return
            
        # 添加程序图标设置（如果有的话）
        try:
            self.window.iconbitmap('icon.ico')  # 如果你有图标文件的话
        except:
            pass
            
        self.window.title("模板文件批量处理工具")
        self.window.geometry("600x400")
        
        # 配置文件路径
        self.config_file = Path.home() / '.template_processor_config.json'
        
        # 初始化变量
        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_path = tk.StringVar()
        
        # 添加替换模式变量
        self.replacement_mode = tk.StringVar(value="dynamic")  # 默认为动态模式
        self.fixed_count = tk.StringVar(value="1")  # 默认固定替换数量为1
        
        # 加载上次的配置
        self.load_config()
        
        self.setup_ui()
        
        # 注册窗口关闭事件
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def load_config(self):
        """加载配置文件"""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.excel_path.set(config.get('excel_path', ''))
                    self.template_path.set(config.get('template_path', ''))
                    self.output_path.set(config.get('output_path', ''))
            except Exception as e:
                self.log(f"加载配置文件失败: {str(e)}")
    
    def save_config(self):
        """保存配置文件"""
        try:
            config = {
                'excel_path': self.excel_path.get(),
                'template_path': self.template_path.get(),
                'output_path': self.output_path.get()
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存配置文件失败: {str(e)}")
    
    def on_closing(self):
        """窗口关闭时保存配置"""
        self.save_config()
        self.window.destroy()
    
    def setup_ui(self):
        # 创建文件选择框架
        file_frame = ttk.LabelFrame(self.window, text="文件选择", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)
        
        # Excel文件选择
        ttk.Label(file_frame, text="Excel文件:").grid(row=0, column=0, sticky="w")
        ttk.Entry(file_frame, textvariable=self.excel_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="浏览", command=self.select_excel).grid(row=0, column=2)
        
        # 模板文件选择
        ttk.Label(file_frame, text="模板文件:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(file_frame, textvariable=self.template_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="浏览", command=self.select_template).grid(row=1, column=2, pady=5)
        
        # 输出路径选择
        ttk.Label(file_frame, text="保存位置:").grid(row=2, column=0, sticky="w")
        ttk.Entry(file_frame, textvariable=self.output_path, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(file_frame, text="浏览", command=self.select_output).grid(row=2, column=2)
        
        # 进度条
        self.progress = ttk.Progressbar(self.window, length=580, mode='determinate')
        self.progress.pack(pady=20, padx=10)
        
        # 添加替换模式选择框架
        mode_frame = ttk.LabelFrame(self.window, text="替换模式", padding=10)
        mode_frame.pack(fill="x", padx=10, pady=5)
        
        # 动态模式选项
        ttk.Radiobutton(
            mode_frame, 
            text="动态替换（根据模板中的标记数量）", 
            variable=self.replacement_mode,
            value="dynamic",
            command=self.toggle_fixed_count
        ).grid(row=0, column=0, sticky="w", padx=5)
        
        # 固定数量模式选项
        ttk.Radiobutton(
            mode_frame, 
            text="固定数量替换", 
            variable=self.replacement_mode,
            value="fixed",
            command=self.toggle_fixed_count
        ).grid(row=1, column=0, sticky="w", padx=5)
        
        # 固定数量输入框
        self.fixed_count_frame = ttk.Frame(mode_frame)
        self.fixed_count_frame.grid(row=1, column=1, sticky="w", padx=5)
        ttk.Label(self.fixed_count_frame, text="每个文件替换行数:").pack(side="left")
        self.fixed_count_entry = ttk.Entry(self.fixed_count_frame, textvariable=self.fixed_count, width=5)
        self.fixed_count_entry.pack(side="left", padx=5)
        
        # 初始化UI状态
        self.toggle_fixed_count()
        
        # 处理按钮
        ttk.Button(self.window, text="开始处理", command=self.process_files).pack(pady=10)
        
        # 日志显示
        log_frame = ttk.LabelFrame(self.window, text="处理日志", padding=10)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.log_text = tk.Text(log_frame, height=10)
        self.log_text.pack(fill="both", expand=True)
        
    def select_excel(self):
        # 获取上次的目录路径
        initial_dir = os.path.dirname(self.excel_path.get()) if self.excel_path.get() else os.path.expanduser("~")
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.excel_path.set(filename)
            self.save_config()
        
    def select_template(self):
        initial_dir = os.path.dirname(self.template_path.get()) if self.template_path.get() else os.path.expanduser("~")
        filename = filedialog.askopenfilename(
            title="选择模板文件",
            initialdir=initial_dir
        )
        if filename:
            self.template_path.set(filename)
            self.save_config()
        
    def select_output(self):
        initial_dir = self.output_path.get() if self.output_path.get() else os.path.expanduser("~")
        directory = filedialog.askdirectory(
            title="选择保存位置",
            initialdir=initial_dir
        )
        if directory:
            self.output_path.set(directory)
            self.save_config()
        
    def log(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        
    def toggle_fixed_count(self):
        """切换固定数量输入框的启用状态"""
        if self.replacement_mode.get() == "fixed":
            self.fixed_count_entry.config(state="normal")
        else:
            self.fixed_count_entry.config(state="disabled")
    
    def process_files(self):
        # 清空之前的状态
        self.progress["value"] = 0
        self.log_text.delete(1.0, tk.END)
        
        if not self.excel_path.get() or not self.template_path.get() or not self.output_path.get():
            messagebox.showerror("错误", "请选择所有必需的文件和路径")
            return
            
        try:
            # 每次都重新读取Excel文件
            self.log("开始读取Excel文件...")
            try:
                df = pd.read_excel(self.excel_path.get(), header=None)
                self.log(f"Excel文件读取成功，共 {len(df)} 行")
            except Exception as e:
                self.log(f"Excel文件读取失败: {str(e)}")
                messagebox.showerror("错误", f"Excel文件读取失败: {str(e)}")
                return
            
            total_rows = len(df)
            self.progress["maximum"] = total_rows
            
            # 每次都重新读取模板文件
            self.log("读取模板文件...")
            try:
                with open(self.template_path.get(), 'r', encoding='utf-8') as template_file:
                    template_content = template_file.read()
                self.log("模板文件读取成功")
            except Exception as e:
                self.log(f"模板文件读取失败: {str(e)}")
                messagebox.showerror("错误", f"模板文件读取失败: {str(e)}")
                return
            
            # 确定每个文件处理的行数
            if self.replacement_mode.get() == "fixed":
                try:
                    replacements_per_file = int(self.fixed_count.get())
                    if replacements_per_file <= 0:
                        raise ValueError("替换行数必须大于0")
                except ValueError as e:
                    self.log(f"固定替换行数设置无效: {str(e)}")
                    messagebox.showerror("错误", "请输入有效的替换行数（必须为大于0的整数）")
                    return
            else:
                # 动态模式：使用模板中标记的数量
                text_count = template_content.count("文案")
                image_count = template_content.count("图片")
                replacements_per_file = min(text_count, image_count)
                if replacements_per_file == 0:
                    messagebox.showerror("错误", "模板文件中未找到'文案'或'图片'标记")
                    return
            
            # 添加状态列和时间列（如果不存在）
            try:
                if len(df.columns) < 5:
                    df[3] = ''  # 状态列
                    df[4] = ''  # 时间列
                self.log("Excel列检查完成")
            except Exception as e:
                self.log(f"Excel列处理失败: {str(e)}")
                messagebox.showerror("错误", f"Excel列处理失败: {str(e)}")
                return
            
            current_row = 0
            file_number = 1
            
            # 确保输出目录存在
            try:
                output_dir = Path(self.output_path.get())
                if not output_dir.exists():
                    output_dir.mkdir(parents=True)
                self.log("输出目录检查完成")
            except Exception as e:
                self.log(f"输出目录创建失败: {str(e)}")
                messagebox.showerror("错误", f"输出目录创建失败: {str(e)}")
                return
            
            # 修改文件命名逻辑
            def generate_output_filename(row_number):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                return f"{timestamp}_{row_number}.txt"  # 时间戳在前，行号在后
            
            while current_row < total_rows:
                current_content = template_content
                rows_processed = 0
                last_processed_row = 0  # 添加变量跟踪最后处理的行号
                
                self.log(f"\n开始处理第{file_number}个文件")
                
                while rows_processed < replacements_per_file and current_row < total_rows:
                    try:
                        row = df.iloc[current_row]
                        
                        # 检查当前行是否已处理
                        if str(row[3]).strip() == '是':
                            self.log(f"跳过已处理的行 {current_row + 1}")
                            current_row += 1
                            continue
                        
                        # 处理文案内容，如果为空则使用空字符串
                        text_content = str(row[1])  # 第二列是文案
                        if pd.isna(row[1]) or text_content.strip() == '':
                            text_content = ''
                            self.log(f"行 {current_row + 1}: 文案内容为空，将替换为空字符串")
                        
                        image_path = str(row[2])    # 第三列是图片
                        
                        self.log(f"处理行 {current_row + 1}: 文案长度={len(text_content)}, 图片路径长度={len(image_path)}")
                        
                        # 查找并替换第一个未替换的"文案"和"图片"
                        if "文案" in current_content:
                            current_content = current_content.replace("文案", text_content, 1)
                        if "图片" in current_content:
                            current_content = current_content.replace("图片", image_path, 1)
                        
                        # 标记处理成功
                        df.at[current_row, 3] = '是'
                        df.at[current_row, 4] = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                        
                        self.log(f"替换成功: 行{current_row + 1}")
                        
                        # 更新最后处理的行号
                        last_processed_row = current_row + 1  # +1 转换为Excel中的行号（从1开始）
                        rows_processed += 1
                        current_row += 1
                        
                    except Exception as e:
                        self.log(f"处理行{current_row + 1}时出错: {str(e)}")
                        self.log(f"错误详情: {repr(e)}")
                        current_row += 1
                        continue
                
                # 检查是否有成功处理的行
                if rows_processed > 0:
                    try:
                        # 使用新的文件命名规则
                        output_file = os.path.join(output_dir, generate_output_filename(file_number))
                        
                        # 保存当前文件
                        with open(output_file, 'w', encoding='utf-8') as new_file:
                            new_file.write(current_content)
                        
                        self.log(f"已生成文件: {output_file}")
                        file_number += 1
                    except Exception as e:
                        self.log(f"保存文件失败: {str(e)}")
                        messagebox.showerror("错误", f"保存文件失败: {str(e)}")
                else:
                    self.log("没有新的数据需要处理")
                    break
                
                # 更新进度条
                self.progress["value"] = current_row
                self.window.update()
                
                # 检查是否还有未处理的行
                remaining_rows = df[df[3] != '是'].shape[0]
                if remaining_rows == 0:
                    self.log("所有数据已处理完成")
                    break
            
            try:
                # 保存Excel文件（覆盖原文件）
                df.to_excel(self.excel_path.get(), index=False, header=False)
                self.log(f"\n已更新Excel文件")
            except Exception as e:
                self.log(f"保存Excel文件失败: {str(e)}")
                messagebox.showerror("错误", f"保存Excel文件失败: {str(e)}")
            
            messagebox.showinfo("完成", f"处理完成！共生成了{file_number-1}个文件。")
            
        except Exception as e:
            self.log(f"发生错误: {str(e)}")
            self.log(f"错误类型: {type(e)}")
            self.log(f"错误详情: {repr(e)}")
            messagebox.showerror("错误", f"处理过程中出现错误：{str(e)}")
            
    def find_all_positions(self, text, target):
        """查找目标字符串在文本中的所有位置"""
        positions = []
        start = 0
        while True:
            pos = text.find(target, start)
            if pos == -1:  # 没有找到更多匹配
                break
            positions.append(pos)
            start = pos + 1
        return positions
    
    def run(self):
        try:
            self.window.mainloop()
        except Exception as e:
            print(f"程序运行错误: {str(e)}")

if __name__ == "__main__":
    try:
        app = TemplateProcessor()
        app.run()
    except Exception as e:
        print(f"程序启动错误: {str(e)}")
        input("按回车键退出...")  # 防止程序立即关闭 