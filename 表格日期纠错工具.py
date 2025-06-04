import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import re
from datetime import datetime
import os
import string

class DateFormatterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("表格日期纠错工具")
        self.root.geometry("800x850")   
        style = ttk.Style()
        style.configure("Title.TLabel", font=("微软雅黑", 11, "bold"))
        style.configure("Info.TLabel", font=("微软雅黑", 9))
        style.configure("Header.TLabel", font=("微软雅黑", 9))
        main_frame = ttk.Frame(root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(title_frame, text="表格日期纠错工具", 
                 style="Title.TLabel").pack(side="left")
        ttk.Label(title_frame, text="支持批量处理多列日期格式", 
                 style="Info.TLabel").pack(side="right")
        file_frame = ttk.LabelFrame(main_frame, text="第一步：选择Excel文件", padding="5")
        file_frame.pack(fill="x", pady=(0, 5))
        self.file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(file_frame, text="浏览文件", command=self.select_file, width=10).pack(side="left", padx=2)
        ttk.Button(file_frame, text="使用说明", command=self.show_help, width=10).pack(side="right", padx=2)
        steps_container = ttk.Frame(main_frame)
        steps_container.pack(fill="both", expand=True, pady=(0, 5))
        sheet_frame = ttk.LabelFrame(steps_container, text="第二步：选择工作表", padding="5")
        sheet_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        self.sheet_var = tk.StringVar()
        self.sheet_combobox = ttk.Combobox(sheet_frame, textvariable=self.sheet_var, state="readonly", height=15)
        self.sheet_combobox.pack(fill="x", padx=5, pady=5)
        self.sheet_combobox.bind('<<ComboboxSelected>>', self.on_sheet_selected)
        columns_frame = ttk.LabelFrame(steps_container, text="第三步：选择要处理的日期列", padding="5")
        columns_frame.pack(side="left", fill="both", expand=True, padx=5)
        columns_container = ttk.Frame(columns_frame)
        columns_container.pack(fill="both", expand=True)
        left_column_frame = ttk.Frame(columns_container)
        left_column_frame.pack(side="left", fill="both", expand=True, padx=(0, 2))
        ttk.Label(left_column_frame, text="可选列：").pack(anchor="w")
        available_list_frame = ttk.Frame(left_column_frame)
        available_list_frame.pack(fill="both", expand=True)
        scrollbar_left = ttk.Scrollbar(available_list_frame)
        scrollbar_left.pack(side="right", fill="y")
        self.available_columns = tk.Listbox(available_list_frame, selectmode="multiple", 
                                       font=("微软雅黑", 9),
                                       height=10,
                                       activestyle="none",
                                       yscrollcommand=scrollbar_left.set)
        self.available_columns.pack(side="left", fill="both", expand=True)
        scrollbar_left.config(command=self.available_columns.yview)
        button_frame = ttk.Frame(columns_container)
        button_frame.pack(side="left", padx=5)
        ttk.Button(button_frame, text="→", command=self.move_to_selected, width=3).pack(pady=2)
        ttk.Button(button_frame, text="←", command=self.move_to_available, width=3).pack(pady=2)
        ttk.Button(button_frame, text="↑", command=lambda: self.move_selected_item(-1), width=3).pack(pady=2)
        ttk.Button(button_frame, text="↓", command=lambda: self.move_selected_item(1), width=3).pack(pady=2)
        right_column_frame = ttk.Frame(columns_container)
        right_column_frame.pack(side="left", fill="both", expand=True, padx=(2, 0))
        ttk.Label(right_column_frame, text="已选列：").pack(anchor="w")
        selected_list_frame = ttk.Frame(right_column_frame)
        selected_list_frame.pack(fill="both", expand=True)
        scrollbar_right = ttk.Scrollbar(selected_list_frame)
        scrollbar_right.pack(side="right", fill="y")
        self.selected_columns = tk.Listbox(selected_list_frame, selectmode="multiple",
                                      font=("微软雅黑", 9),
                                      height=10,
                                      activestyle="none",
                                      yscrollcommand=scrollbar_right.set)
        self.selected_columns.pack(side="left", fill="both", expand=True)
        scrollbar_right.config(command=self.selected_columns.yview)
        format_frame = ttk.LabelFrame(steps_container, text="第四步：选择目标格式", padding="5")
        format_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))
        format_container = ttk.Frame(format_frame)
        format_container.pack(fill="both", expand=True, pady=5)
        self.target_format = tk.StringVar(value="xxxx.xx.xx")
        formats = [
            ("年月日格式：xxxx.xx.xx", "xxxx.xx.xx"),
            ("年月日格式：xxxx-xx-xx", "xxxx-xx-xx"),
            ("年月格式：xxxx.xx", "xxxx.xx"),
            ("年月格式：xxxx-xx", "xxxx-xx")
        ]
        for text, value in formats:
            ttk.Radiobutton(format_container, text=text, value=value,
                           variable=self.target_format).pack(anchor="w", pady=5)
        button_frame = ttk.LabelFrame(main_frame, text="第五步：执行转换", padding="5")
        button_frame.pack(fill="x", pady=(0, 5))
        
        ttk.Button(button_frame, text="开始格式化", 
                  command=self.format_dates,
                  width=20).pack(pady=5)
        preview_frame = ttk.LabelFrame(main_frame, text="数据预览（选中列的前15行）", padding="5")
        preview_frame.pack(fill="both", expand=True)
        preview_container = ttk.Frame(preview_frame)
        preview_container.pack(fill="both", expand=True)
        tree_scroll_y = ttk.Scrollbar(preview_container)
        tree_scroll_y.pack(side="right", fill="y")
        tree_scroll_x = ttk.Scrollbar(preview_container, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")
        self.preview_tree = ttk.Treeview(preview_container,
                                       yscrollcommand=tree_scroll_y.set,
                                       xscrollcommand=tree_scroll_x.set,
                                       height=15) 
        self.preview_tree.pack(fill="both", expand=True)
        tree_scroll_y.config(command=self.preview_tree.yview)
        tree_scroll_x.config(command=self.preview_tree.xview)
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill="x", pady=(5, 0))
        
        self.status_var = tk.StringVar(value="请选择Excel文件")
        ttk.Label(status_frame, text="状态：").pack(side="left")
        self.status_bar = ttk.Label(status_frame, textvariable=self.status_var,
                                  relief="sunken", padding=(5, 2))
        self.status_bar.pack(side="left", fill="x", expand=True)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, 
                                          mode='determinate',
                                          variable=self.progress_var,
                                          maximum=100)
        self.progress_bar.pack(fill="x", pady=(5, 0))
        
        self.df = None
        self.excel_file = None
        self.column_stats = {}
        
    def on_sheet_selected(self, event):
        if self.excel_file is not None:
            sheet_name = self.sheet_var.get()
            self.show_processing(True, "正在切换工作表...")
            self.root.after(100, lambda: self._load_selected_sheet(sheet_name))

    def _load_selected_sheet(self, sheet_name):
        try:
            self.load_sheet(sheet_name)
        finally:
            self.show_processing(False)
            self.status_var.set(f"工作表 {sheet_name} 加载完成")

    def load_sheet(self, sheet_name):
        try:
            self.show_processing(True)
            self.df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
            self.available_columns.delete(0, tk.END)
            self.selected_columns.delete(0, tk.END)
            for idx, col in enumerate(self.df.columns):
                col_letter = self.get_column_letter(idx)
                self.available_columns.insert(tk.END, f"{col_letter} | {col}")
            self.preview_tree["columns"] = ()
            self.status_var.set(f"已加载工作表：{sheet_name}")
            
        except Exception as e:
            messagebox.showerror("错误", f"无法加载工作表: {str(e)}")
            self.status_var.set("工作表加载失败")
        finally:
            self.show_processing(False)
            
    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path.set(file_path)
            try:
                self.show_processing(True, "正在加载Excel文件...")
                self.root.after(100, lambda: self._load_file(file_path))
            except Exception as e:
                messagebox.showerror("错误", f"无法加载文件: {str(e)}")
                self.status_var.set("文件加载失败")
                self.show_processing(False)

    def _load_file(self, file_path):
        try:
            self.excel_file = file_path
            xl = pd.ExcelFile(file_path)
            sheet_names = xl.sheet_names
            self.sheet_combobox['values'] = sheet_names
            if sheet_names:
                self.sheet_combobox.set(sheet_names[0])
                self.root.after(100, lambda: self._load_first_sheet(sheet_names[0]))
        except Exception as e:
            messagebox.showerror("错误", f"无法加载文件: {str(e)}")
            self.status_var.set("文件加载失败")
            self.show_processing(False)

    def _load_first_sheet(self, sheet_name):
        try:
            self.load_sheet(sheet_name)
        finally:
            self.show_processing(False)
            self.status_var.set("文件加载完成")

    def show_processing(self, show=True, message=None):
        if show:
            self.progress_var.set(0)
            self._animate_progress()
            if message:
                self.status_var.set(message)
        else:
            if self.progress_var.get() < 100:
                self.progress_var.set(100)
            self.root.after(500, self._stop_progress)

    def _animate_progress(self):
        current = self.progress_var.get()
        if current < 100:
            self.progress_var.set(min(current + 2, 99))
            self.root.after(20, self._animate_progress)

    def _stop_progress(self):
        self.progress_var.set(0)

    def show_help(self):
        help_window = tk.Toplevel(self.root)
        help_window.title("使用说明")
        help_window.geometry("500x600")
        
        help_text = """
功能说明：
本工具用于批量纠正Excel文件中的日期格式。

主要功能：
1. 多列同时处理
2. 自动识别多种日期格式
3. 生成错误报告
4. 智能日期格式识别和验证

支持的输入：
1. 纯数字格式：
   8位数字：20230405（年月日）
   6位数字：230405（年月日）

2. 带分隔符格式：
   xxxx.xx.xx（年.月.日）
   xxxx-xx-xx（年-月-日）
   xxxx/xx/xx（年/月/日）
   xx.xx.xx（年.月.日）
   xx-xx-xx（年-月-日）
   xx/xx/xx（年/月/日）

3. 年月格式：
   xxxx.xx（年.月）
   xxxx-xx（年-月）
   xxxx/xx（年/月）

4. 特殊格式：
   xxxx.xxxx（年.月日）

智能处理规则：
1. 自动补全两位年份：
   23 -> 2023
   13 -> 2013

2. 自动补全个位数字：
   2023.1.1 -> 2023.01.01
   2023-9-9 -> 2023-09-09

3. 数据验证：
   月份范围检查（1-12）
   日期范围检查（1-31）
   年份格式检查（2位或4位）

转换结果：
将根据所选格式转换为标准格式：
1. xxxx.xx.xx（年.月.日）
2. xxxx-xx-xx（年-月-日）
3. xxxx.xx（年.月）
4. xxxx-xx（年-月）

使用步骤：
1. 点击"选择文件"按钮选择Excel文件
2. 在列表中选择需要处理的列（可多选）
3. 选择目标日期格式
4. 点击"格式化日期"按钮
5. 查看处理结果和错误报告

注意事项：
  空白单元格将保持不变
  无法识别的格式会在报告中列出
  每个有问题的列会生成单独的错误报告
  日期格式识别会自动移除多余的空格
  支持多种分隔符（点号、横杠、斜杠）



                                by：孙政
"""
        help_text_widget = scrolledtext.ScrolledText(help_window, wrap=tk.WORD)
        help_text_widget.pack(fill="both", expand=True, padx=10, pady=10)
        help_text_widget.insert(tk.END, help_text)
        help_text_widget.configure(state='disabled')
        ttk.Button(help_window, text="关闭", 
                  command=help_window.destroy).pack(pady=10)
        help_window.transient(self.root)
        help_window.grab_set()
        self.root.wait_window(help_window)
        
    def get_column_letter(self, index):
        result = ""
        while index >= 0:
            result = string.ascii_uppercase[index % 26] + result
            index = index // 26 - 1
        return result
        
    def update_preview(self):
        self.show_processing(True, "正在更新预览...")
        self.root.after(100, self._do_update_preview)

    def _do_update_preview(self):
        try:
            for item in self.preview_tree.get_children():
                self.preview_tree.delete(item)
                
            if self.df is None or self.df.empty:
                return
                
            selected_columns = [item.split(" | ")[1] for item in self.selected_columns.get(0, tk.END)]
            if not selected_columns:
                self.preview_tree["columns"] = ()
                return
            self.preview_tree["columns"] = selected_columns
            self.preview_tree["show"] = "headings"
            for idx, col in enumerate(selected_columns):
                col_letter = self.get_column_letter(idx)
                display_name = f"{col_letter} | {col}"
                self.preview_tree.heading(col, text=display_name)
                sample_data = self.df[col].head(15).astype(str)
                max_width = max(
                    len(display_name),
                    sample_data.str.len().max() if not sample_data.empty else 0
                ) * 10
                self.preview_tree.column(col, width=min(max_width, 300), minwidth=50)
            preview_data = self.df[selected_columns].head(15)
            for idx, row in preview_data.iterrows():
                values = []
                for val in row:
                    if pd.isna(val) or val == "":
                        values.append("（此项为空白）")
                    else:
                        values.append(str(val))
                self.preview_tree.insert("", "end", values=values)
        finally:
            self.show_processing(False)
            self.status_var.set("预览更新完成")

    def format_date(self, date_str, stats, target_format):
        stats['total'] += 1
        
        if pd.isna(date_str) or date_str == "":
            stats['empty'] += 1
            stats['empty_rows'].append(self.current_row_idx)
            return date_str
            
        if not isinstance(date_str, str):
            date_str = str(date_str)
            
        date_str = date_str.strip()
        
        try:
            year = None
            month = None
            day = None
            
            date_str = date_str.replace(" ", "")
            
            # 1. 纯数字
            if re.match(r'^\d{6,8}$', date_str):
                if len(date_str) == 8: 
                    year = date_str[:4]
                    month = date_str[4:6]
                    day = date_str[6:]
                elif len(date_str) == 6:  
                    year = '20' + date_str[:2]
                    month = date_str[2:4]
                    day = date_str[4:]
            
            # 2.  xxxx.xxxx 
            elif re.match(r'^\d{4}\.\d{4}$', date_str):
                year = date_str[:4]
                month = date_str[5:7]
                day = date_str[7:]
                
            # 3. 带分隔符
            elif re.search(r'[./-]', date_str):
                parts = re.split(r'[./-]', date_str)
                if len(parts) >= 2:
                    # 年
                    year = parts[0]
                    if len(year) == 2:
                        year = '20' + year
                    elif len(year) != 4:
                        raise ValueError("年份格式不正确")
                    
                    # 月
                    month = parts[1]
                    if len(month) == 1:
                        month = '0' + month
                    elif len(month) > 2:
                        raise ValueError("月份格式不正确")
                    
                    # 日
                    if len(parts) > 2:
                        day = parts[2]
                        if len(day) == 1:
                            day = '0' + day
                        elif len(day) > 2:
                            raise ValueError("日期格式不正确")
            
            if month and (not month.isdigit() or int(month) < 1 or int(month) > 12):
                raise ValueError("月份无效")
            if day and (not day.isdigit() or int(day) < 1 or int(day) > 31):
                raise ValueError("日期无效")
            
            if year and month:
                if target_format in ["xxxx.xx.xx", "xxxx-xx-xx"]:
                    if not day:
                        raise ValueError("需要日期部分")
                    separator = '.' if target_format == "xxxx.xx.xx" else '-'
                    result = f"{year}{separator}{month}{separator}{day}"
                elif target_format in ["xxxx.xx", "xxxx-xx"]:
                    separator = '.' if target_format == "xxxx.xx" else '-'
                    result = f"{year}{separator}{month}"
                else:
                    raise ValueError("不支持的目标格式")
                    
                stats['processed'] += 1
                return result
                
            raise ValueError("无法解析日期格式")
            
        except Exception as e:
            if date_str.strip():
                stats['unrecognized'] += 1
                stats['unrecognized_data'].append({
                    'row': self.current_row_idx + 2,
                    'value': date_str,
                    'error': str(e)
                })
            return date_str
            
    def generate_error_report(self, original_file_path, column_name, column_index, stats):
        try:
            dir_path = os.path.dirname(original_file_path)
            file_name = os.path.splitext(os.path.basename(original_file_path))[0]
            column_display_name = f"{self.get_column_letter(column_index)}_{column_name}"
            safe_column_name = "".join(c for c in column_display_name if c.isalnum() or c in ('_', '-'))
            error_file = os.path.join(dir_path, f"{file_name}_{safe_column_name}_错误报告.xlsx")
            with pd.ExcelWriter(error_file, engine='openpyxl') as writer:
                stats_df = pd.DataFrame([{
                    '列名': column_name,
                    '总行数': stats['total'],
                    '成功处理': stats['processed'],
                    '空白单元格': stats['empty'],
                    '无法识别': stats['unrecognized']
                }])
                stats_df.to_excel(writer, sheet_name='处理统计', index=False)
                if stats['unrecognized_data']:
                    unrecognized_df = pd.DataFrame(stats['unrecognized_data'])
                    unrecognized_df.columns = ['Excel行号', '原始值', '错误描述']
                    unrecognized_df.to_excel(writer, sheet_name='无法识别的日期', index=False)
                if stats['empty_rows']:
                    empty_df = pd.DataFrame({
                        'Excel行号': [idx + 2 for idx in stats['empty_rows']]
                    })
                    empty_df.to_excel(writer, sheet_name='空白单元格', index=False)
            
            return error_file
            
        except Exception as e:
            self.status_var.set(f"生成错误报告失败: {column_name}")
            messagebox.showerror("错误", f"生成错误报告时出错: {str(e)}")
            return None

    def format_dates(self):
        if self.df is None:
            messagebox.showerror("错误", "请先选择文件")
            return
            
        selected_columns = list(self.selected_columns.get(0, tk.END))
        if not selected_columns:
            messagebox.showerror("错误", "请选择要格式化的列")
            return
            
        self.show_processing(True, "正在格式化日期...")
        self.root.after(100, lambda: self._do_format_dates(selected_columns))

    def _do_format_dates(self, selected_columns):
        try:
            selected_columns = [col.split(" | ")[1] for col in selected_columns]
            error_files = []
            total_columns = len(selected_columns)
            
            for col_idx, column in enumerate(selected_columns, 1):
                self.status_var.set(f"正在处理第 {col_idx}/{total_columns} 列: {column}")
                self.root.update()
                self.column_stats[column] = {
                    'total': 0,
                    'processed': 0,
                    'empty': 0,
                    'unrecognized': 0,
                    'empty_rows': [],
                    'unrecognized_data': []
                }
                for idx, value in enumerate(self.df[column]):
                    self.current_row_idx = idx
                    self.df.at[idx, column] = self.format_date(
                        value, 
                        self.column_stats[column],
                        self.target_format.get()
                    )
            
            self._save_processed_file(error_files)
            
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出错: {str(e)}")
            self.status_var.set("格式化失败")
        finally:
            self.show_processing(False)

    def _save_processed_file(self, error_files):
        try:
            original_path = self.file_path.get()
            dir_path = os.path.dirname(original_path)
            file_name = os.path.splitext(os.path.basename(original_path))[0]
            
            default_save_path = os.path.join(dir_path, f"{file_name}_已处理.xlsx")
            
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=os.path.basename(default_save_path),
                initialdir=dir_path
            )
            
            if file_path:
                self.status_var.set("正在保存文件...")
                self.df.to_excel(file_path, index=False)
                
                successful_reports = []
                for col_idx, column in enumerate(self.selected_columns.get(0, tk.END)):
                    col_name = column.split(" | ")[1]
                    stats = self.column_stats[col_name]
                    
                    if stats['unrecognized'] > 0 or stats['empty'] > 0:
                        error_file = self.generate_error_report(file_path, col_name, col_idx, stats)
                        if error_file:
                            successful_reports.append(error_file)
                
                self._show_processing_results(successful_reports)
                self.update_preview()
                
        except Exception as e:
            messagebox.showerror("错误", f"保存文件时出错: {str(e)}")
            self.status_var.set("保存失败")

    def _show_processing_results(self, error_files):
        try:
            result_message = "处理完成！\n\n"
            
            # 汇总所有列的统计信息
            total_stats = {
                'total': 0,
                'processed': 0,
                'empty': 0,
                'unrecognized': 0
            }
            
            for column in self.selected_columns.get(0, tk.END):
                col_name = column.split(" | ")[1]
                stats = self.column_stats[col_name]
                for key in total_stats:
                    total_stats[key] += stats[key]
            
            result_message += f"总计处理: {total_stats['total']} 行\n"
            result_message += f"成功转换: {total_stats['processed']} 行\n"
            result_message += f"空白单元格: {total_stats['empty']} 行\n"
            result_message += f"无法识别: {total_stats['unrecognized']} 行\n\n"
            
            result_message += f"处理后的文件已保存在:\n{os.path.dirname(self.file_path.get())}\n"
            
            if error_files:
                result_message += f"\n已生成 {len(error_files)} 个错误报告文件。"
            
            messagebox.showinfo("处理完成", result_message)
            self.status_var.set("处理完成")
            
        except Exception as e:
            messagebox.showerror("错误", f"显示处理结果时出错: {str(e)}")
            self.status_var.set("显示结果失败")

    def move_to_selected(self):
        self.show_processing(True, "正在更新列表...")
        self.root.after(100, self._do_move_to_selected)

    def _do_move_to_selected(self):
        try:
            selections = self.available_columns.curselection()
            for i in reversed(selections):
                value = self.available_columns.get(i)
                self.selected_columns.insert(tk.END, value)
                self.available_columns.delete(i)
            self.update_preview()
        finally:
            self.show_processing(False)
            self.status_var.set("列表更新完成")

    def move_to_available(self):
        self.show_processing(True, "正在更新列表...")
        self.root.after(100, self._do_move_to_available)

    def _do_move_to_available(self):
        try:
            selections = self.selected_columns.curselection()
            for i in reversed(selections):
                value = self.selected_columns.get(i)
                self.available_columns.insert(tk.END, value)
                self.selected_columns.delete(i)
            self.update_preview()
        finally:
            self.show_processing(False)
            self.status_var.set("列表更新完成")

    def move_selected_item(self, direction):
        selections = self.selected_columns.curselection()
        if not selections:
            return
            
        index = selections[0]
        if direction == -1 and index == 0: 
            return
        if direction == 1 and index == self.selected_columns.size() - 1: 
            return
            
        value = self.selected_columns.get(index)
        self.selected_columns.delete(index)
        new_index = index + direction
        self.selected_columns.insert(new_index, value)
        self.selected_columns.selection_set(new_index)
        self.update_preview()

if __name__ == "__main__":
    root = tk.Tk()
    app = DateFormatterApp(root)
    root.mainloop()
