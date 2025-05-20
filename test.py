import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import json
import re
import sys

class ExtractColumn:
    """提取列类，表示要从匹配文件中提取的列配置"""
    def __init__(self, name, search_names, enabled=True, is_primary=False):
        self.name = name  # 列的显示名称
        self.search_names = search_names  # 搜索列名列表
        self.enabled = enabled  # 是否启用此列
        self.is_primary = is_primary  # 是否为主键列（用于比对）
    
    def to_dict(self):
        """将列配置转换为字典，用于JSON序列化"""
        return {
            "name": self.name,
            "search_names": self.search_names,
            "enabled": self.enabled,
            "is_primary": self.is_primary
        }
    
    @classmethod
    def from_dict(cls, column_dict):
        """从字典创建列配置对象"""
        return cls(
            name=column_dict.get("name", ""),
            search_names=column_dict.get("search_names", []),
            enabled=column_dict.get("enabled", True),
            is_primary=column_dict.get("is_primary", False)
        )

class ComparisonRule:
    """比对规则类，用于存储和应用自定义比对规则"""
    def __init__(self, name, conditions, match_all=True, enabled=True):
        self.name = name  # 规则名称
        self.conditions = conditions  # 条件列表
        self.match_all = match_all  # 是否需要满足所有条件
        self.enabled = enabled  # 是否启用此规则
    
    def to_dict(self):
        """将规则转换为字典，用于JSON序列化"""
        return {
            "name": self.name,
            "conditions": [condition.to_dict() for condition in self.conditions],
            "match_all": self.match_all,
            "enabled": self.enabled
        }
    
    @classmethod
    def from_dict(cls, rule_dict):
        """从字典创建规则对象"""
        conditions = [ColumnCondition.from_dict(cond) for cond in rule_dict.get("conditions", [])]
        return cls(
            name=rule_dict.get("name", "未命名规则"),
            conditions=conditions,
            match_all=rule_dict.get("match_all", True),
            enabled=rule_dict.get("enabled", True)
        )
    
    def match(self, row, columns_map):
        """检查行是否符合规则条件"""
        if not self.enabled or not self.conditions:
            return False
        
        # 检查每个条件
        matches = []
        for condition in self.conditions:
            # 获取实际列名
            actual_column = columns_map.get(condition.column_name.lower().strip())
            if not actual_column:
                # 尝试部分匹配
                for col_lower, col in columns_map.items():
                    if condition.column_name.lower().strip() in col_lower or col_lower in condition.column_name.lower().strip():
                        actual_column = col
                        break
            
            if not actual_column:
                # 列不存在
                matches.append(False)
                continue
            
            # 获取单元格值
            cell_value = str(row[actual_column]).strip()
            if not condition.case_sensitive:
                cell_value = cell_value.lower()
            
            # 检查是否匹配
            condition_match = False
            for value in condition.search_values:
                search_value = value if condition.case_sensitive else value.lower()
                
                if condition.is_regex:
                    try:
                        if re.search(search_value, cell_value):
                            condition_match = True
                            break
                    except re.error:
                        # 正则表达式错误，视为不匹配
                        pass
                elif condition.exact_match:
                    # 精确匹配
                    if cell_value == search_value:
                        condition_match = True
                        break
                else:
                    # 部分匹配
                    if search_value in cell_value:
                        condition_match = True
                        break
            
            matches.append(condition_match)
        
        # 根据match_all判断最终结果
        if self.match_all:
            return all(matches)
        else:
            return any(matches)

class ColumnCondition:
    """列条件类，表示对单个列的匹配条件"""
    def __init__(self, column_name, search_values, case_sensitive=False, is_regex=False, exact_match=False):
        self.column_name = column_name  # 列名
        self.search_values = search_values  # 搜索值列表
        self.case_sensitive = case_sensitive  # 是否区分大小写
        self.is_regex = is_regex  # 是否使用正则表达式
        self.exact_match = exact_match  # 是否精确匹配
    
    def to_dict(self):
        """将条件转换为字典，用于JSON序列化"""
        return {
            "column_name": self.column_name,
            "search_values": self.search_values,
            "case_sensitive": self.case_sensitive,
            "is_regex": self.is_regex,
            "exact_match": self.exact_match
        }
    
    @classmethod
    def from_dict(cls, condition_dict):
        """从字典创建条件对象"""
        return cls(
            column_name=condition_dict.get("column_name", ""),
            search_values=condition_dict.get("search_values", []),
            case_sensitive=condition_dict.get("case_sensitive", False),
            is_regex=condition_dict.get("is_regex", False),
            exact_match=condition_dict.get("exact_match", False)
        )

class ExcelComparator:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel文件比对工具")
        
        # 设置窗口最小尺寸，确保所有内容能够显示
        self.root.minsize(900, 650)
        
        # 初始窗口大小
        self.root.geometry("1000x700")
        
        # 设置窗口图标（如果有图标文件）
        # try:
        #    self.root.iconbitmap("icon.ico")
        # except:
        #    pass
        
        # 应用样式主题
        self.setup_styles()
        
        # 设置变量
        self.master_file_path = tk.StringVar()
        self.folder_path = tk.StringVar()
        self.result_data = []
        self.master_sheet_name = None  # 保存选择的工作表名称
        
        # 添加模型匹配模式变量，默认为完全匹配
        self.exact_model_match = tk.BooleanVar(value=True)
        
        # 添加无规则时提取所有行的设置，默认为False
        self.extract_all_when_no_rules = tk.BooleanVar(value=False)
        
        # 比对规则列表
        self.comparison_rules = []
        
        # 添加提取列配置列表
        self.extract_columns = []
        
        # 配置文件路径
        try:
            # PyInstaller打包后的应用程序位置
            if getattr(sys, 'frozen', False):
                application_path = os.path.dirname(sys.executable)
            else:
                application_path = os.path.dirname(os.path.abspath(__file__))
            self.config_file = os.path.join(application_path, "config.json")
            print(f"配置文件路径: {self.config_file}")
        except Exception as e:
            print(f"配置文件路径设置错误: {str(e)}")
            # 备用方案：保存在当前工作目录
            self.config_file = "config.json"
        
        # 加载上次的设置和规则
        self.load_settings()
        
        # 如果没有任何规则，创建默认规则
        if not self.comparison_rules:
            self.create_default_rules()
        
        # 如果没有任何提取列配置，创建默认配置
        if not self.extract_columns:
            self.create_default_extract_columns()
        
        # 创建界面
        self.create_widgets()
        
    def setup_styles(self):
        """设置应用样式"""
        style = ttk.Style()
        
        # 使用默认主题，避免样式问题
        try:
            style.theme_use('vista')  # 在Windows上通常效果较好
        except:
            try:
                style.theme_use('winnative')  # 备选主题
            except:
                pass  # 如果都失败，使用系统默认主题
        
        # 自定义按钮样式
        style.configure('TButton', font=('Microsoft YaHei UI', 10))
        style.configure('Primary.TButton', font=('Microsoft YaHei UI', 10, 'bold'))
        
        # 避免设置可能导致问题的背景颜色
        # style.map('Primary.TButton', background=[('active', '#0069d9'), ('pressed', '#005cbf')])
        
        # 自定义标签样式
        style.configure('TLabel', font=('Microsoft YaHei UI', 10))
        style.configure('Header.TLabel', font=('Microsoft YaHei UI', 12, 'bold'))
        
        # 自定义框架样式
        style.configure('Card.TFrame', relief='ridge', borderwidth=1)
        
        # 自定义Treeview样式
        style.configure("Treeview", 
                        font=('Microsoft YaHei UI', 9),
                        rowheight=25)
        style.configure("Treeview.Heading", 
                        font=('Microsoft YaHei UI', 10, 'bold'))
        
        # 添加交替行颜色
        style.map('Treeview', 
                  background=[('selected', '#3399ff')],
                  foreground=[('selected', 'white')])
        
        # 设置奇数行和偶数行的背景色
        self.tree_tag_configure = True  # 标记是否已配置树形视图标签
    
    def create_widgets(self):
        # 创建主框架 - 使用网格布局
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 顶部标题
        header_label = ttk.Label(main_frame, text="Excel文件比对工具", style='Header.TLabel')
        header_label.pack(pady=(0, 10))
        
        # 文件选择区域 - 使用卡片式设计
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 主文件选择框
        master_frame = ttk.Frame(file_frame)
        master_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(master_frame, text="总Excel文件:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(master_frame, textvariable=self.master_file_path, width=60).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        button_frame = ttk.Frame(master_frame)
        button_frame.pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="浏览...", command=self.browse_master_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="选择工作表", command=self.select_worksheet).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="预览", command=self.preview_file).pack(side=tk.LEFT, padx=2)
        
        # 文件夹选择框
        folder_frame = ttk.Frame(file_frame)
        folder_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(folder_frame, text="比对文件夹:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(folder_frame, textvariable=self.folder_path, width=60).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(folder_frame, text="浏览...", command=self.browse_folder).pack(side=tk.LEFT, padx=5)
        
        # 添加模型匹配模式选项 - 放在新的框架中
        match_mode_frame = ttk.Frame(file_frame)
        match_mode_frame.pack(fill=tk.X, pady=5)
        
        ttk.Checkbutton(match_mode_frame, text="完全匹配模型名（精确匹配Model）", 
                        variable=self.exact_model_match).pack(side=tk.LEFT, padx=(0, 5))
        
        # 在match_mode_frame中添加无规则时提取所有行的选项
        ttk.Checkbutton(match_mode_frame, text="无规则时提取所有行", 
                        variable=self.extract_all_when_no_rules).pack(side=tk.LEFT, padx=(20, 5))
        
        # 动作按钮框架
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 执行按钮 - 使用突出显示的样式
        ttk.Button(action_frame, text="开始比对", command=self.compare_files, style='Primary.TButton').pack(side=tk.LEFT, padx=5)
        
        # 管理规则按钮
        ttk.Button(action_frame, text="管理比对规则", command=self.manage_rules).pack(side=tk.LEFT, padx=5)
        
        # 在action_frame中添加管理提取列按钮
        ttk.Button(action_frame, text="管理提取列", command=self.manage_extract_columns).pack(side=tk.LEFT, padx=5)
        
        # 添加使用说明按钮 - 移到管理提取列按钮后面
        ttk.Button(action_frame, text="使用说明", command=self.show_help).pack(side=tk.LEFT, padx=5)
        
        # 结果框架
        result_frame = ttk.LabelFrame(main_frame, text="比对结果", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建Treeview用于显示结果 - 增加列宽以确保内容显示
        primary_column_name = "对应文件Part No"  # 默认名称
        
        # 查找启用的主键列
        for column in self.extract_columns:
            if column.is_primary and column.enabled:
                primary_column_name = f"对应文件{column.name}"
                break
        
        # 设置结果列名
        columns = ("序号", "Model", "总文件Part No", primary_column_name, "比对结果")
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings")
        
        # 设置列标题和宽度
        column_widths = {
            "序号": 60,
            "Model": 120,
            "总文件Part No": 200,
            "比对结果": 100
        }
        # 为主键列动态添加宽度
        column_widths[primary_column_name] = 200
        
        for col in columns:
            self.result_tree.heading(col, text=col)
            self.result_tree.column(col, width=column_widths[col], minwidth=50)
        
        # 添加水平和垂直滚动条
        v_scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        self.result_tree.configure(xscrollcommand=h_scrollbar.set)
        
        # 使用网格布局放置树形视图和滚动条
        self.result_tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        # 设置行列权重使得树形视图可随窗口调整大小
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        # 底部工具栏
        toolbar_frame = ttk.Frame(main_frame, padding="5")
        toolbar_frame.pack(fill=tk.X, pady=5)
        
        # 状态标签 - 左侧
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(toolbar_frame, textvariable=self.status_var, anchor=tk.W)
        status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 功能按钮 - 右侧
        button_frame = ttk.Frame(toolbar_frame)
        button_frame.pack(side=tk.RIGHT)
        
        ttk.Button(button_frame, text="导出结果", command=self.export_results).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="合并匹配文件", command=self.merge_matched_files).pack(side=tk.LEFT, padx=5)
    
    def browse_master_file(self):
        """浏览并选择主Excel文件"""
        # 修改文件类型列表，添加CSV文件选项
        file_types = [("所有支持的文件", "*.xlsx *.xls *.csv"), 
                      ("Excel文件", "*.xlsx *.xls"), 
                      ("CSV文件", "*.csv"), 
                      ("所有文件", "*.*")]
        
        file_path = filedialog.askopenfilename(title="选择主文件", filetypes=file_types)
        if file_path:
            self.master_file_path.set(file_path)
            self.master_sheet_name = None  # 重置工作表选择
            self.save_settings()
    
    def browse_folder(self):
        folder_path = filedialog.askdirectory(
            title="选择比对文件夹",
            initialdir=self.folder_path.get() if self.folder_path.get() else None
        )
        if folder_path:
            self.folder_path.set(folder_path)
            self.save_settings()
    
    def select_worksheet(self):
        """打开工作表选择对话框"""
        file_path = self.master_file_path.get()
        if not file_path:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
            
        # 检查文件是否存在
        if not os.path.exists(file_path):
            messagebox.showerror("错误", f"找不到文件: {file_path}")
            return
            
        # 检查文件类型
        if not file_path.lower().endswith(('.xlsx', '.xls')):
            messagebox.showinfo("提示", "只有Excel文件才需要选择工作表")
            return
            
        try:
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names
            
            if not sheet_names:
                messagebox.showinfo("提示", "此Excel文件不包含任何工作表")
                return
                
            # 初始化result_vars字典
            result_vars = {
                "confirmed": False,
                "model": "",
                "partno": ""
            }
                
            # 创建工作表选择对话框
            select_dialog = tk.Toplevel(self.root)
            select_dialog.title("选择工作表")
            select_dialog.geometry("300x400")
            select_dialog.transient(self.root)
            select_dialog.grab_set()
            
            ttk.Label(select_dialog, text="请选择要使用的工作表:").pack(pady=10)
            
            # 工作表列表框
            sheet_frame = ttk.Frame(select_dialog)
            sheet_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
            
            # 添加滚动条
            scrollbar = ttk.Scrollbar(sheet_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            sheet_listbox = tk.Listbox(sheet_frame, yscrollcommand=scrollbar.set)
            sheet_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=sheet_listbox.yview)
            
            # 填充工作表列表
            for sheet in sheet_names:
                sheet_listbox.insert(tk.END, sheet)
            
            # 如果之前已选择过工作表，则预选该工作表
            if self.master_sheet_name and self.master_sheet_name in sheet_names:
                idx = sheet_names.index(self.master_sheet_name)
                sheet_listbox.selection_set(idx)
                sheet_listbox.see(idx)
            else:
                # 默认选择第一个
                sheet_listbox.selection_set(0)
            
            # 确认按钮
            def on_confirm():
                selected_indices = sheet_listbox.curselection()
                if selected_indices:
                    self.master_sheet_name = sheet_names[selected_indices[0]]
                    result_vars["confirmed"] = True
                    # 读取工作表获取列信息
                    try:
                        df = pd.read_excel(file_path, sheet_name=self.master_sheet_name)
                        columns = df.columns.tolist()
                        # 尝试找到Model和Part No列
                        for col in columns:
                            col_lower = str(col).lower()
                            if "model" in col_lower or "型号" in col_lower:
                                result_vars["model"] = col
                            if "part" in col_lower or "零件" in col_lower or "料号" in col_lower:
                                result_vars["partno"] = col
                    except Exception as e:
                        print(f"读取工作表列信息时出错: {str(e)}")
                    
                    messagebox.showinfo("成功", f"已选择工作表: {self.master_sheet_name}")
                    select_dialog.destroy()
                else:
                    messagebox.showwarning("警告", "请选择一个工作表")
            
            ttk.Button(select_dialog, text="确认", command=on_confirm).pack(pady=10)
            
            # 等待对话框关闭
            self.root.wait_window(select_dialog)
            
            # 如果用户取消了选择
            if not result_vars["confirmed"]:
                self.update_status("就绪")
                return
                    
            # 使用用户选择的列
            model_col = result_vars["model"]
            partno_col = result_vars["partno"]
            
            if model_col or partno_col:
                print(f"用户选择 - Model列: {model_col}, Part No列: {partno_col}")
        
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel文件失败: {str(e)}")
    
    def save_settings(self):
        """保存当前设置和规则到配置文件"""
        settings = {
            "master_file_path": self.master_file_path.get(),
            "folder_path": self.folder_path.get(),
            "rules": [rule.to_dict() for rule in self.comparison_rules],
            "extract_columns": [column.to_dict() for column in self.extract_columns],
            "master_sheet_name": self.master_sheet_name,
            "exact_model_match": self.exact_model_match.get(),
            "extract_all_when_no_rules": self.extract_all_when_no_rules.get()
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存设置时出错: {e}")
    
    def load_settings(self):
        """从配置文件加载设置和规则"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.master_file_path.set(settings.get("master_file_path", ""))
                    self.folder_path.set(settings.get("folder_path", ""))
                    self.master_sheet_name = settings.get("master_sheet_name")
                    
                    # 加载模型匹配模式设置，默认为True（完全匹配）
                    self.exact_model_match.set(settings.get("exact_model_match", True))
                    
                    # 加载无规则时提取所有行设置，默认为False
                    self.extract_all_when_no_rules.set(settings.get("extract_all_when_no_rules", False))
                    
                    # 加载规则
                    rules_data = settings.get("rules", [])
                    self.comparison_rules = [ComparisonRule.from_dict(rule_dict) for rule_dict in rules_data]
                    
                    # 加载提取列配置
                    extract_columns_data = settings.get("extract_columns", [])
                    self.extract_columns = [ExtractColumn.from_dict(column_dict) for column_dict in extract_columns_data]
        except Exception as e:
            print(f"加载设置时出错: {e}")
            # 出错时使用空字符串作为默认值
            self.master_file_path.set("")
            self.folder_path.set("")
            self.comparison_rules = []
            self.extract_columns = []
    
    def create_default_rules(self):
        """创建默认的比对规则"""
        # PCB Assembly + Source Right/Left 规则
        pcb_rule = ComparisonRule(
            name="PCB Assembly 规则",
            conditions=[
                ColumnCondition("Item Desc", ["PCB Assembly"], False, False),
                ColumnCondition("Item Spec", ["Source Right", "Source Left"], False, False)
            ],
            match_all=True,
            enabled=True
        )
        self.comparison_rules.append(pcb_rule)
    
    def manage_rules(self):
        """打开规则管理对话框"""
        # 创建一个新窗口
        rules_dialog = tk.Toplevel(self.root)
        rules_dialog.title("比对规则管理")
        rules_dialog.minsize(850, 600)  # 设置最小尺寸
        rules_dialog.geometry("950x650")  # 设置初始尺寸
        
        # 创建主框架并使用padding
        main_frame = ttk.Frame(rules_dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 添加标题
        ttk.Label(main_frame, text="比对规则管理", style='Header.TLabel').pack(pady=(0, 10))
        
        # 创建水平分隔的左右面板 - 使用PanedWindow
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 左侧面板 - 规则列表
        list_frame = ttk.LabelFrame(paned, text="规则列表", padding="5")
        
        # 右侧面板 - 规则详情
        detail_frame = ttk.LabelFrame(paned, text="规则详情", padding="5")
        
        # 添加到PanedWindow
        paned.add(list_frame, weight=1)
        paned.add(detail_frame, weight=2)
        
        # 规则列表
        rules_frame = ttk.Frame(list_frame)
        rules_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 添加滚动条
        rules_scrollbar = ttk.Scrollbar(rules_frame)
        rules_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        rules_listbox = tk.Listbox(rules_frame, height=10, width=30, yscrollcommand=rules_scrollbar.set)
        rules_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        rules_scrollbar.config(command=rules_listbox.yview)
        
        # 规则详情区域
        # 规则属性框架
        rule_props_frame = ttk.Frame(detail_frame)
        rule_props_frame.pack(fill=tk.X, pady=5)
        
        # 标题和输入框使用网格布局使其对齐
        ttk.Label(rule_props_frame, text="规则名称:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        rule_name_var = tk.StringVar()
        ttk.Entry(rule_props_frame, textvariable=rule_name_var, width=30).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 复选框
        match_all_var = tk.BooleanVar(value=True)
        enabled_var = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(rule_props_frame, text="需要满足所有条件", variable=match_all_var).grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Checkbutton(rule_props_frame, text="启用此规则", variable=enabled_var).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 条件列表框架
        conditions_outer_frame = ttk.LabelFrame(detail_frame, text="条件列表")
        conditions_outer_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 使用内部框架并添加滚动条
        conditions_frame = ttk.Frame(conditions_outer_frame)
        conditions_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        cond_scrollbar = ttk.Scrollbar(conditions_frame)
        cond_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        conditions_listbox = tk.Listbox(conditions_frame, height=6, width=50, yscrollcommand=cond_scrollbar.set)
        conditions_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        cond_scrollbar.config(command=conditions_listbox.yview)
        
        # 条件编辑框架
        condition_edit_frame = ttk.LabelFrame(detail_frame, text="条件编辑")
        condition_edit_frame.pack(fill=tk.X, pady=5)
        
        # 使用网格布局使控件对齐
        ttk.Label(condition_edit_frame, text="列名:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        column_name_var = tk.StringVar()
        ttk.Entry(condition_edit_frame, textvariable=column_name_var, width=20).grid(row=0, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        
        ttk.Label(condition_edit_frame, text="搜索值:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        search_value_var = tk.StringVar()
        ttk.Entry(condition_edit_frame, textvariable=search_value_var, width=30).grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        
        # 复选框区域
        checkbox_frame = ttk.Frame(condition_edit_frame)
        checkbox_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # 添加复选框
        case_sensitive_var = tk.BooleanVar(value=False)
        is_regex_var = tk.BooleanVar(value=False)
        exact_match_var = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(checkbox_frame, text="区分大小写", variable=case_sensitive_var).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Checkbutton(checkbox_frame, text="正则表达式", variable=is_regex_var).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Checkbutton(checkbox_frame, text="精确匹配", variable=exact_match_var).pack(side=tk.LEFT)
        
        # 条件按钮框架
        condition_buttons_frame = ttk.Frame(detail_frame)
        condition_buttons_frame.pack(fill=tk.X, pady=5)
        
        # 规则和条件的索引变量
        current_rule_index = {"value": -1}
        current_condition_index = {"value": -1}
        
        # 当前规则和条件数据
        current_conditions = []
        
        # 加载所有规则到列表
        for rule in self.comparison_rules:
            rules_listbox.insert(tk.END, rule.name)
        
        # 更新条件列表
        def update_conditions_list():
            conditions_listbox.delete(0, tk.END)
            for condition in current_conditions:
                conditions_listbox.insert(tk.END, f"{condition.column_name}: {', '.join(condition.search_values)}")
        
        # 加载规则到编辑区域
        def load_rule(index):
            if 0 <= index < len(self.comparison_rules):
                rule = self.comparison_rules[index]
                current_rule_index["value"] = index
                
                rule_name_var.set(rule.name)
                match_all_var.set(rule.match_all)
                enabled_var.set(rule.enabled)
                
                # 加载条件
                current_conditions.clear()
                for condition in rule.conditions:
                    current_conditions.append(condition)
                
                update_conditions_list()
                
                # 清除当前条件选择
                current_condition_index["value"] = -1
                column_name_var.set("")
                search_value_var.set("")
                case_sensitive_var.set(False)
                is_regex_var.set(False)
                exact_match_var.set(False)
        
        def load_condition(index):
            """加载条件到编辑区域"""
            if 0 <= index < len(current_conditions):
                condition = current_conditions[index]
                current_condition_index["value"] = index
                
                column_name_var.set(condition.column_name)
                search_value_var.set(', '.join(condition.search_values))
                case_sensitive_var.set(condition.case_sensitive)
                is_regex_var.set(condition.is_regex)
                exact_match_var.set(condition.exact_match)
        
        def on_rule_select(evt):
            """当选择规则时"""
            selection = rules_listbox.curselection()
            if selection:
                index = selection[0]
                load_rule(index)
        
        def on_condition_select(evt):
            """当选择条件时"""
            selection = conditions_listbox.curselection()
            if selection:
                index = selection[0]
                load_condition(index)
        
        # 规则操作函数
        def add_rule():
            """添加新规则"""
            new_rule = ComparisonRule(
                name="新规则",
                conditions=[],
                match_all=True,
                enabled=True
            )
            self.comparison_rules.append(new_rule)
            
            # 更新列表
            rules_listbox.insert(tk.END, new_rule.name)
            
            # 选择新规则
            index = len(self.comparison_rules) - 1
            rules_listbox.selection_clear(0, tk.END)
            rules_listbox.selection_set(index)
            rules_listbox.see(index)
            load_rule(index)
        
        def update_rule():
            """更新当前规则"""
            index = current_rule_index["value"]
            if 0 <= index < len(self.comparison_rules):
                rule = self.comparison_rules[index]
                rule.name = rule_name_var.get()
                rule.match_all = match_all_var.get()
                rule.enabled = enabled_var.get()
                
                # 更新条件
                rule.conditions = current_conditions.copy()
                
                # 更新列表显示
                rules_listbox.delete(index)
                rules_listbox.insert(index, rule.name)
                rules_listbox.selection_set(index)
        
        def delete_rule():
            """删除当前规则"""
            index = current_rule_index["value"]
            if 0 <= index < len(self.comparison_rules):
                # 删除规则
                del self.comparison_rules[index]
                
                # 更新列表
                rules_listbox.delete(index)
                
                # 选择其他规则
                if self.comparison_rules:
                    new_index = min(index, len(self.comparison_rules) - 1)
                    rules_listbox.selection_set(new_index)
                    load_rule(new_index)
                else:
                    # 没有规则了，清空编辑区
                    rule_name_var.set("")
                    match_all_var.set(True)
                    enabled_var.set(True)
                    current_conditions.clear()
                    update_conditions_list()
                    current_rule_index["value"] = -1
        
        def add_condition():
            """添加新条件"""
            new_condition = ColumnCondition(
                column_name="",
                search_values=[""],
                case_sensitive=False,
                is_regex=False,
                exact_match=False
            )
            current_conditions.append(new_condition)
            
            # 更新列表
            conditions_listbox.insert(tk.END, f"{new_condition.column_name}: {', '.join(new_condition.search_values)}")
            
            # 选择新条件
            index = len(current_conditions) - 1
            conditions_listbox.selection_clear(0, tk.END)
            conditions_listbox.selection_set(index)
            conditions_listbox.see(index)
            load_condition(index)
        
        def update_condition():
            """更新当前条件"""
            index = current_condition_index["value"]
            if 0 <= index < len(current_conditions):
                condition = current_conditions[index]
                condition.column_name = column_name_var.get()
                # 处理搜索值，支持中英文逗号
                search_value_text = search_value_var.get()
                # 先替换中文逗号为英文逗号，然后分割
                search_value_text = search_value_text.replace('，', ',')
                search_values = [v.strip() for v in search_value_text.split(',') if v.strip()]
                
                if search_values:
                    condition.search_values = search_values
                condition.case_sensitive = case_sensitive_var.get()
                condition.is_regex = is_regex_var.get()
                condition.exact_match = exact_match_var.get()
                
                # 更新列表显示
                conditions_listbox.delete(index)
                conditions_listbox.insert(index, f"{condition.column_name}: {', '.join(condition.search_values)}")
                conditions_listbox.selection_set(index)
        
        def delete_condition():
            """删除当前条件"""
            index = current_condition_index["value"]
            if 0 <= index < len(current_conditions):
                del current_conditions[index]
                
                # 更新列表
                conditions_listbox.delete(index)
                
                # 选择其他条件
                if current_conditions:
                    new_index = min(index, len(current_conditions) - 1)
                    conditions_listbox.selection_set(new_index)
                    load_condition(new_index)
                else:
                    # 没有条件了，清空编辑区
                    column_name_var.set("")
                    search_value_var.set("")
                    case_sensitive_var.set(False)
                    is_regex_var.set(False)
                    exact_match_var.set(False)
                    current_condition_index["value"] = -1
        
        # 绑定选择事件
        rules_listbox.bind('<<ListboxSelect>>', on_rule_select)
        conditions_listbox.bind('<<ListboxSelect>>', on_condition_select)
        
        # 规则按钮
        rule_buttons_frame = ttk.Frame(list_frame)
        rule_buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(rule_buttons_frame, text="添加规则", command=add_rule).pack(side=tk.LEFT, padx=5)
        ttk.Button(rule_buttons_frame, text="更新规则", command=update_rule).pack(side=tk.LEFT, padx=5)
        ttk.Button(rule_buttons_frame, text="删除规则", command=delete_rule).pack(side=tk.LEFT, padx=5)
        
        # 条件按钮
        ttk.Button(condition_buttons_frame, text="添加条件", command=add_condition).pack(side=tk.LEFT, padx=5)
        ttk.Button(condition_buttons_frame, text="更新条件", command=update_condition).pack(side=tk.LEFT, padx=5)
        ttk.Button(condition_buttons_frame, text="删除条件", command=delete_condition).pack(side=tk.LEFT, padx=5)
        
        # 确认和取消按钮
        buttons_frame = ttk.Frame(rules_dialog)
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)
        
        def on_save():
            # 如果有正在编辑的规则，先保存它
            if current_rule_index["value"] >= 0:
                update_rule()
            self.save_settings()
            rules_dialog.destroy()
        
        ttk.Button(buttons_frame, text="保存并关闭", command=on_save).pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="取消", command=rules_dialog.destroy).pack(side=tk.RIGHT, padx=5)
        
        # 如果有规则，默认选择第一个
        if self.comparison_rules:
            rules_listbox.selection_set(0)
            load_rule(0)
        
        # 使窗口在父窗口中居中
        rules_dialog.transient(self.root)
        rules_dialog.update_idletasks()
        width = rules_dialog.winfo_width()
        height = rules_dialog.winfo_height()
        x = self.root.winfo_x() + (self.root.winfo_width() - width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - height) // 2
        rules_dialog.geometry(f"{width}x{height}+{x}+{y}")
        
        # 设置为模态对话框
        rules_dialog.grab_set()
        rules_dialog.focus_set()
        rules_dialog.wait_window()
    
    def read_file(self, file_path):
        """根据文件类型读取文件内容"""
        if file_path.lower().endswith(('.xlsx', '.xls')):
            # 读取Excel文件
            if self.master_sheet_name and file_path == self.master_file_path.get():
                # 如果是主文件并且已选择工作表，则使用选择的工作表
                return pd.read_excel(file_path, sheet_name=self.master_sheet_name)
            else:
                # 尝试获取所有工作表名称
                xls = pd.ExcelFile(file_path)
                sheet_names = xls.sheet_names
                
                if not sheet_names:
                    raise ValueError(f"Excel文件不包含任何工作表: {file_path}")
                
                # 尝试读取所有工作表，直到找到有数据的工作表
                for sheet in sheet_names:
                    temp_df = pd.read_excel(file_path, sheet_name=sheet)
                    if not temp_df.empty:
                        print(f"在工作表 '{sheet}' 中找到数据")
                        return temp_df
                
                # 如果所有工作表都为空，尝试不同的header选项
                for sheet in sheet_names:
                    for header_row in range(5):  # 尝试前5行作为表头
                        try:
                            temp_df = pd.read_excel(file_path, sheet_name=sheet, header=header_row)
                            if not temp_df.empty and len(temp_df.columns) > 1:  # 确保有多列数据
                                return temp_df
                        except Exception:
                            pass
                
                raise ValueError(f"无法从Excel文件中读取有效数据: {file_path}")
        
        elif file_path.lower().endswith('.csv'):
            # 读取CSV文件，优先尝试中文编码
            # 根据日志分析，调整编码尝试顺序，将中文编码放在前面
            encodings = ['gb18030', 'gbk', 'gb2312', 'utf-8', 'latin1']
            
            for encoding in encodings:
                try:
                    print(f"尝试使用编码 {encoding} 读取CSV文件...")
                    df = pd.read_csv(file_path, encoding=encoding)
                    if not df.empty:
                        print(f"成功使用编码 {encoding} 读取CSV文件")
                        return df
                except Exception as e:
                    print(f"尝试使用编码 {encoding} 读取失败: {str(e)}")
            
            # 如果所有编码都失败，询问用户
            encoding_dialog = tk.Toplevel(self.root)
            encoding_dialog.title("选择编码")
            encoding_dialog.geometry("350x200")
            encoding_dialog.transient(self.root)
            encoding_dialog.grab_set()
            
            ttk.Label(encoding_dialog, text="无法自动识别CSV文件编码，请选择:").pack(pady=10)
            
            encoding_var = tk.StringVar(value="utf-8")
            encoding_combo = ttk.Combobox(encoding_dialog, textvariable=encoding_var)
            encoding_combo['values'] = ('utf-8', 'gb18030', 'gbk', 'gb2312', 'latin1', 'utf-16', 'ascii')
            encoding_combo.pack(pady=10)
            
            result = {"df": None, "success": False}
            
            def on_confirm():
                try:
                    result["df"] = pd.read_csv(file_path, encoding=encoding_var.get())
                    result["success"] = True
                    encoding_dialog.destroy()
                except Exception as e:
                    messagebox.showerror("错误", f"使用编码 {encoding_var.get()} 读取失败: {str(e)}")
            
            ttk.Button(encoding_dialog, text="确认", command=on_confirm).pack(pady=10)
            
            # 等待对话框关闭
            self.root.wait_window(encoding_dialog)
            
            if result["success"] and result["df"] is not None:
                return result["df"]
            
            raise ValueError(f"无法读取CSV文件: {file_path}")
        else:
            raise ValueError(f"不支持的文件类型: {file_path}")
    
    def compare_files(self):
        # 更新状态
        self.update_status("正在准备比对...")
        
        master_path = self.master_file_path.get()
        folder_path = self.folder_path.get()
        
        if not master_path or not folder_path:
            messagebox.showerror("错误", "请选择总Excel文件和比对文件夹")
            self.update_status("就绪")
            return
        
        # 创建输出文件夹
        output_folder = os.path.join(folder_path, "匹配文件")
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            
        try:
            # 清空先前的结果
            for item in self.result_tree.get_children():
                self.result_tree.delete(item)
            self.result_data = []
            
            # 检查文件是否存在和可访问
            if not os.path.exists(master_path):
                messagebox.showerror("错误", f"找不到文件: {master_path}")
                self.update_status("就绪")
                return
            
            self.update_status("正在读取主文件...")
            
            try:
                # 读取主文件
                master_df = self.read_file(master_path)
                
                # 显示读取到的数据基本信息
                print(f"成功读取数据，行数: {len(master_df)}")
                print(f"列名: {master_df.columns.tolist()}")
                print("前3行内容:")
                print(master_df.head(3))
                
            except Exception as e:
                import traceback
                error_details = traceback.format_exc()
                messagebox.showerror("错误", f"读取文件失败: {str(e)}\n\n请确保文件格式正确且未被其他程序锁定。\n\n详细信息:\n{error_details}")
                self.update_status("就绪")
                return
            
            # 列名规范化处理
            # 将DataFrame的列名转换为小写，并存储原始列名与小写列名的映射
            columns_lower = {str(col).lower().strip(): col for col in master_df.columns}
            
            # 显示处理后的列名映射，帮助调试
            print("处理后的列名映射:")
            for k, v in columns_lower.items():
                print(f"  {k} -> {v}")
            
            # 检查必要的列是否存在（不区分大小写）
            model_col = None
            partno_col = None
            
            # 查找model列（尝试多种可能的写法）
            for possible_name in ['model', 'model no', 'model number', 'model#', 'models', '型号', '模型', 'model号']:
                possible_lower = possible_name.lower()
                # 精确匹配
                if possible_lower in columns_lower:
                    model_col = columns_lower[possible_lower]
                    print(f"找到model列(精确匹配): {model_col}")
                    break
                # 部分匹配
                for col_lower, col in columns_lower.items():
                    if possible_lower in col_lower or col_lower in possible_lower:
                        model_col = col
                        print(f"找到model列(部分匹配): {model_col}")
                        break
                if model_col:
                    break
            
            # 查找Part No列（尝试多种可能的写法）
            for possible_name in ['part no', 'partno', 'part number', 'part#', 'partnumber', 'part_no', 'part-no', 'part', '零件号', '零件编号', '料号']:
                possible_lower = possible_name.lower()
                # 尝试精确匹配
                if possible_lower in columns_lower:
                    partno_col = columns_lower[possible_lower]
                    print(f"找到part no列(精确匹配): {partno_col}")
                    break
                # 尝试部分匹配
                for col_lower, col in columns_lower.items():
                    if possible_lower in col_lower or col_lower in possible_lower:
                        partno_col = col
                        print(f"找到part no列(部分匹配): {partno_col}")
                        break
                if partno_col:
                    break
            
            # 如果找不到列，让用户手动选择
            if not model_col or not partno_col:
                # 创建列选择对话框
                select_dialog = tk.Toplevel(self.root)
                select_dialog.title("列选择")
                select_dialog.geometry("400x300")
                select_dialog.transient(self.root)
                select_dialog.grab_set()
                
                ttk.Label(select_dialog, text="无法自动识别必要的列，请手动选择:").pack(pady=10)
                
                # Model列选择
                model_frame = ttk.Frame(select_dialog)
                model_frame.pack(fill=tk.X, padx=10, pady=5)
                ttk.Label(model_frame, text="选择Model列:").pack(side=tk.LEFT)
                model_var = tk.StringVar()
                model_combo = ttk.Combobox(model_frame, textvariable=model_var, values=list(master_df.columns))
                model_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
                
                # Part No列选择
                partno_frame = ttk.Frame(select_dialog)
                partno_frame.pack(fill=tk.X, padx=10, pady=5)
                ttk.Label(partno_frame, text="选择Part No列:").pack(side=tk.LEFT)
                partno_var = tk.StringVar()
                partno_combo = ttk.Combobox(partno_frame, textvariable=partno_var, values=list(master_df.columns))
                partno_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
                
                # 返回选择结果
                result_vars = {"model": None, "partno": None, "confirmed": False}
                
                def on_confirm():
                    if model_var.get() and partno_var.get():
                        result_vars["model"] = model_var.get()
                        result_vars["partno"] = partno_var.get()
                        result_vars["confirmed"] = True
                        select_dialog.destroy()
                    else:
                        messagebox.showwarning("警告", "请选择两个列")
                
                ttk.Button(select_dialog, text="确认", command=on_confirm).pack(pady=10)
                
                # 等待对话框关闭
                self.root.wait_window(select_dialog)
                
                # 如果用户取消了选择
                if not result_vars["confirmed"]:
                    self.update_status("就绪")
                    return
                    
                # 使用用户选择的列
                model_col = result_vars["model"]
                partno_col = result_vars["partno"]
                
                print(f"用户选择 - Model列: {model_col}, Part No列: {partno_col}")
            
            # 遍历主表中的每一行
            total_rows = len(master_df)
            for index, row in master_df.iterrows():
                # 更新进度状态
                self.update_status(f"正在比对第 {index+1}/{total_rows} 行...")
                
                model = str(row[model_col]).strip()
                master_part_no = str(row[partno_col]).strip()
                
                # 查找对应文件
                model_file = None
                for file in os.listdir(folder_path):
                    if file.lower().endswith(('.xlsx', '.xls', '.csv')):
                        # 根据匹配模式选择不同的比对逻辑
                        if self.exact_model_match.get():
                            # 完全匹配模式 - 使用正则表达式匹配完整模型名称
                            pattern = r'(^|[^\w])' + re.escape(model.lower()) + r'([^\w]|$)'
                            if re.search(pattern, file.lower()):
                                model_file = file
                                break
                        else:
                            # 部分匹配模式 - 保持原有逻辑
                            if model.lower() in file.lower():
                                model_file = file
                                break
                
                if model_file:
                    # 读取对应的文件
                    file_path = os.path.join(folder_path, model_file)
                    try:
                        compare_df = self.read_file(file_path)
                        
                        # 新的Part No提取逻辑
                        compare_part_nos = self.extract_special_part_nos(compare_df, file_path, output_folder)
                        
                        if compare_part_nos:
                            # 有匹配的Part No
                            
                            # 修改为显示所有匹配结果
                            # 首先检查是否有完全匹配的结果
                            exact_matches = [pn for pn in compare_part_nos if pn == master_part_no]
                            
                            if exact_matches:
                                # 有完全匹配的结果
                                for match_pn in exact_matches:
                                    result = (index + 1, model, master_part_no, match_pn, "匹配")
                                    self.result_data.append(result)
                                    self.result_tree.insert("", tk.END, values=result)
                                
                                # 同时显示其他不匹配的结果，但标记为"其他结果"
                                other_part_nos = [pn for pn in compare_part_nos if pn != master_part_no]
                                for other_pn in other_part_nos:
                                    result = (index + 1, model, master_part_no, other_pn, "其他结果")
                                    self.result_data.append(result)
                                    self.result_tree.insert("", tk.END, values=result)
                            else:
                                # 没有完全匹配的结果，显示所有结果为"不匹配"
                                for pn in compare_part_nos:
                                    result = (index + 1, model, master_part_no, pn, "不匹配")
                                    self.result_data.append(result)
                                    self.result_tree.insert("", tk.END, values=result)
                        else:
                            result = (index + 1, model, master_part_no, "未找到符合条件的Part No", "不匹配")
                            self.result_data.append(result)
                            self.result_tree.insert("", tk.END, values=result)
                    except Exception as e:
                        result = (index + 1, model, master_part_no, f"文件读取错误: {str(e)}", "错误")
                        self.result_data.append(result)
                        self.result_tree.insert("", tk.END, values=result)
                else:
                    result = (index + 1, model, master_part_no, "未找到对应文件", "错误")
                    self.result_data.append(result)
                    self.result_tree.insert("", tk.END, values=result)
            
            # 给用户提供结果摘要
            matches_count = len([r for r in self.result_data if r[4] == "匹配"])
            non_matches_count = len([r for r in self.result_data if r[4] == "不匹配"])
            other_results_count = len([r for r in self.result_data if r[4] == "其他结果"])
            error_count = len([r for r in self.result_data if r[4] == "错误"])
            
            summary = f"比对完成!\n\n匹配: {matches_count}\n不匹配: {non_matches_count}\n其他结果: {other_results_count}\n错误: {error_count}"
            messagebox.showinfo("完成", summary)
            
            self.update_status("比对完成")
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("错误", f"比对过程中发生错误: {str(e)}\n\n详细信息:\n{error_details}")
            self.update_status("就绪")
    
    def extract_special_part_nos(self, df, file_path, output_folder):
        """
        根据自定义规则和提取列配置从DataFrame中提取数据
        """
        try:
            # 将列名转为小写便于查找
            columns_lower = {str(col).lower().strip(): col for col in df.columns}
            
            # 查找主键列（通常是Part No）
            primary_column = None
            primary_actual_col = None
            
            for column in self.extract_columns:
                if column.is_primary and column.enabled:
                    primary_column = column
                    break
            
            if not primary_column:
                # 如果没有设置主键列，使用第一个启用的列作为主键
                for column in self.extract_columns:
                    if column.enabled:
                        primary_column = column
                        break
            
            if not primary_column:
                print("没有找到可用的主键列配置")
                return []
            
            # 查找主键列的实际列名
            for search_name in primary_column.search_names:
                search_name_lower = search_name.lower().strip()
                if search_name_lower in columns_lower:
                    primary_actual_col = columns_lower[search_name_lower]
                    print(f"找到主键列 '{primary_column.name}': {primary_actual_col}")
                    break
                
                # 尝试部分匹配
                for col_lower, col in columns_lower.items():
                    if search_name_lower in col_lower or col_lower in search_name_lower:
                        primary_actual_col = col
                        print(f"找到主键列 '{primary_column.name}'(部分匹配): {primary_actual_col}")
                        break
                
                if primary_actual_col:
                    break
            
            if not primary_actual_col:
                print(f"找不到主键列 '{primary_column.name}'")
                return []
            
            # 检查是否有启用的规则
            has_enabled_rules = any(rule.enabled for rule in self.comparison_rules)
            has_rule_conditions = any(rule.enabled and rule.conditions for rule in self.comparison_rules)
            
            # 如果没有规则或没有条件，且设置为提取所有行
            if (not has_enabled_rules or not has_rule_conditions) and self.extract_all_when_no_rules.get():
                print("没有启用的规则或规则没有条件，且设置了提取所有行")
                matching_rows = df.to_dict('records')  # 提取所有行
            else:
                # 应用所有启用的规则
                matching_rows = []
                
                # 检查每一行
                for index, row in df.iterrows():
                    # 检查每个规则
                    for rule in self.comparison_rules:
                        if rule.enabled and rule.match(row, columns_lower):
                            matching_rows.append(row)
                            part_no = str(row[primary_actual_col]).strip()
                            print(f"找到匹配行: {primary_column.name}={part_no}, 规则={rule.name}")
                            break  # 一旦找到匹配的规则就不再检查其他规则
            
            # 如果有匹配行，创建新的DataFrame并保存到输出文件夹
            if matching_rows:
                matched_df = pd.DataFrame(matching_rows)
                
                # 生成输出文件名
                file_name = os.path.basename(file_path)
                file_base, file_ext = os.path.splitext(file_name)
                output_file = os.path.join(output_folder, f"{file_base}_匹配{file_ext}")
                
                # 保存文件
                if file_ext.lower() in ['.xlsx', '.xls']:
                    matched_df.to_excel(output_file, index=False)
                else:  # CSV文件
                    matched_df.to_csv(output_file, encoding='gb18030', index=False)
                
                print(f"已保存匹配文件: {output_file}")
                
                # 返回匹配行的主键列值
                return [str(row[primary_actual_col]).strip() for row in matching_rows]
            else:
                print("未找到满足规则的行")
                return []
                
        except Exception as e:
            print(f"提取特殊Part No时出错: {str(e)}")
            import traceback
            traceback.print_exc()
            return []
    
    def export_results(self):
        if not self.result_data:
            messagebox.showinfo("提示", "没有可导出的结果")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if file_path:
            try:
                # 创建DataFrame并导出
                result_df = pd.DataFrame(
                    self.result_data, 
                    columns=["序号", "Model", "总文件Part No", "对应文件Part No", "比对结果"]
                )
                result_df.to_excel(file_path, index=False)
                messagebox.showinfo("成功", f"结果已导出至 {file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出过程中发生错误: {str(e)}")
    
    def merge_matched_files(self):
        """合并所有匹配文件为一个单一文件"""
        folder_path = self.folder_path.get()
        if not folder_path:
            messagebox.showerror("错误", "请先选择比对文件夹")
            return
            
        output_folder = os.path.join(folder_path, "匹配文件")
        if not os.path.exists(output_folder):
            messagebox.showinfo("提示", "匹配文件夹不存在，请先运行比对")
            return
            
        # 获取所有匹配文件
        matched_files = []
        for file in os.listdir(output_folder):
            if "_匹配" in file and file.lower().endswith(('.xlsx', '.xls', '.csv')):
                matched_files.append(os.path.join(output_folder, file))
                
        if not matched_files:
            messagebox.showinfo("提示", "未找到任何匹配文件")
            return
        
        # 弹出对话框询问合并方式
        merge_method = messagebox.askyesnocancel(
            "选择合并方式", 
            "如何合并匹配文件?\n\n是 - 合并为一张表\n否 - 每个文件作为单独的工作表\n取消 - 取消操作"
        )
        
        if merge_method is None:  # 用户点击了取消
            return
            
        # 询问用户保存合并文件的路径
        merged_file_path = filedialog.asksaveasfilename(
            title="保存合并文件",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            initialdir=output_folder,
            initialfile="合并匹配文件.xlsx"
        )
        
        if not merged_file_path:
            return  # 用户取消了操作
        
        # 如果用户选择了每个文件作为单独的工作表，并且输出格式不是Excel，需要提示
        if not merge_method and not merged_file_path.lower().endswith('.xlsx'):
            messagebox.showinfo("提示", "多工作表模式只支持Excel格式，将自动更改为.xlsx")
            merged_file_path = os.path.splitext(merged_file_path)[0] + '.xlsx'
            
        try:
            if merge_method:  # 合并为一张表
                # 读取并合并所有文件
                all_dfs = []
                
                for file_path in matched_files:
                    try:
                        # 读取文件
                        if file_path.lower().endswith(('.xlsx', '.xls')):
                            df = pd.read_excel(file_path)
                        else:  # CSV文件
                            df = pd.read_csv(file_path, encoding='gb18030')
                        
                        # 添加来源文件名列
                        file_name = os.path.basename(file_path)
                        df['来源文件'] = file_name
                        
                        all_dfs.append(df)
                    except Exception as e:
                        print(f"读取文件 {file_path} 时出错: {str(e)}")
                
                if not all_dfs:
                    messagebox.showerror("错误", "无法读取任何匹配文件")
                    return
                    
                # 合并所有DataFrame
                merged_df = pd.concat(all_dfs, ignore_index=True)
                
                # 保存合并后的文件
                if merged_file_path.lower().endswith('.xlsx'):
                    merged_df.to_excel(merged_file_path, index=False)
                else:  # CSV文件
                    merged_df.to_csv(merged_file_path, encoding='gb18030', index=False)
                    
                messagebox.showinfo("成功", f"已成功合并 {len(all_dfs)} 个文件到单个表格!\n保存至: {merged_file_path}")
            
            else:  # 每个文件作为单独的工作表
                # 使用ExcelWriter将每个文件写入不同的工作表
                with pd.ExcelWriter(merged_file_path, engine='openpyxl') as writer:
                    success_count = 0
                    
                    for file_path in matched_files:
                        try:
                            # 读取文件
                            if file_path.lower().endswith(('.xlsx', '.xls')):
                                df = pd.read_excel(file_path)
                            else:  # CSV文件
                                df = pd.read_csv(file_path, encoding='gb18030')
                            
                            # 设置工作表名称 - 使用文件名但去掉扩展名和"_匹配"部分
                            file_name = os.path.basename(file_path)
                            sheet_name = os.path.splitext(file_name)[0]
                            if "_匹配" in sheet_name:
                                sheet_name = sheet_name.replace("_匹配", "")
                                
                            # Excel工作表名称有长度限制
                            if len(sheet_name) > 31:  # Excel限制工作表名为31个字符
                                sheet_name = sheet_name[:31]
                                
                            # 避免重复的工作表名
                            original_name = sheet_name
                            counter = 1
                            while sheet_name in writer.sheets:
                                sheet_name = f"{original_name[:27]}_{counter}"
                                counter += 1
                            
                            # 写入到工作表
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            success_count += 1
                            
                        except Exception as e:
                            print(f"处理文件 {file_path} 时出错: {str(e)}")
                    
                    if success_count == 0:
                        messagebox.showerror("错误", "无法读取任何匹配文件")
                        return
                
                messagebox.showinfo("成功", f"已成功将 {success_count} 个文件合并为独立工作表!\n保存至: {merged_file_path}")
            
            # 询问是否打开合并后的文件
            if messagebox.askyesno("提示", "是否打开合并后的文件?"):
                os.startfile(merged_file_path)
                
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("错误", f"合并文件时发生错误: {str(e)}\n\n详细信息:\n{error_details}")

    def update_status(self, message):
        """更新状态栏消息"""
        self.status_var.set(message)
        self.root.update_idletasks()  # 立即更新UI

    def show_help(self):
        """显示使用说明对话框"""
        help_dialog = tk.Toplevel(self.root)
        help_dialog.title("Excel文件比对工具使用说明")
        help_dialog.minsize(800, 600)
        help_dialog.geometry("900x700")
        
        # 主框架
        main_frame = ttk.Frame(help_dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建一个带滚动条的文本框
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        help_text = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set, padx=10, pady=10)
        help_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=help_text.yview)
        
        # 设置文本样式
        help_text.tag_configure("title", font=("Microsoft YaHei UI", 16, "bold"))
        help_text.tag_configure("subtitle", font=("Microsoft YaHei UI", 14, "bold"))
        help_text.tag_configure("heading", font=("Microsoft YaHei UI", 12, "bold"))
        help_text.tag_configure("subheading", font=("Microsoft YaHei UI", 11, "bold"), foreground="#333333")
        help_text.tag_configure("normal", font=("Microsoft YaHei UI", 10))
        help_text.tag_configure("emphasis", font=("Microsoft YaHei UI", 10, "bold"), foreground="#0066cc")
        help_text.tag_configure("warning", font=("Microsoft YaHei UI", 10), foreground="#cc3300")
        help_text.tag_configure("code", font=("Consolas", 9), background="#f5f5f5")
        
        # 添加使用说明内容
        help_content = """Excel文件比对工具 - 详细使用说明

一、软件概述
===============================

本工具专为解决大量数据比对需求而设计，能够快速比较主文件和多个其他文件中的数据。工具通过自定义规则识别和提取符合特定条件的数据，适用于各种数据管理场景：

• 清单比对：验证不同版本清单中的数据差异
• 物料管理：识别特定种类的物料记录
• 数据提取：从大量文件中提取符合条件的记录
• 文件整合：将分散在多个文件中的匹配数据合并处理

二、工作原理与逻辑流程
===============================

1. 核心比对逻辑
--------------------------
• 读取主文件中的每一行数据
• 根据Model字段，在比对文件夹中查找相关文件
• 在找到的文件中，应用自定义规则寻找匹配行
• 根据匹配情况，将结果分类保存和显示
• 提取匹配行保存为单独文件，供后续处理

2. 规则引擎工作方式
--------------------------
• 规则定义了"寻找什么"和"如何匹配"
• 每条规则包含一个或多个条件
• 条件定义了针对特定列的匹配要求
• 支持简单匹配、精确匹配和正则表达式
• 可设置需满足全部条件或任一条件

3. 文件处理逻辑
--------------------------
• 自动识别Excel和CSV文件格式
• 智能检测表头和关键列
• 支持多种编码格式的CSV文件
• 处理结果保存在"匹配文件"子文件夹
• 提供多种结果合并方式

三、详细操作流程
===============================

1. 准备工作
--------------------------
• 确保主文件(总文件)格式规范，包含必要的Model列和Part No列
• 整理比对文件夹，放入所有需要比对的文件
• 思考并规划所需的比对规则

2. 设置文件路径
--------------------------
步骤1: 点击"浏览..."按钮选择主Excel文件
步骤2: 如需必要，点击"选择工作表"选择特定工作表
步骤3: 点击"浏览..."按钮选择包含比对文件的文件夹
       (所有设置会自动保存，下次启动时自动恢复)

3. 设置匹配选项
--------------------------
• 完全匹配模型名：启用时使用精确匹配逻辑寻找Model对应文件
• 无规则时提取所有行：启用时，当没有定义比对规则或规则无条件时，将提取文件中的所有行用于比对，否则将返回"未找到符合条件的Part No"

4. 配置比对规则
--------------------------
步骤1: 点击"管理比对规则"按钮打开规则管理窗口
步骤2: 点击"添加规则"创建新规则并命名
步骤3: 设置是否需要满足所有条件，及是否启用规则
步骤4: 点击"添加条件"添加匹配条件
步骤5: 配置条件的列名、搜索值和匹配方式
       - 列名：输入要查找的表头名称(支持模糊匹配)
       - 搜索值：输入要匹配的内容(多个值用逗号分隔)
       - 选择适当的匹配选项：
         * 区分大小写：是否区分字母大小写
         * 正则表达式：使用正则表达式进行复杂匹配
         * 精确匹配：要求内容完全一致，而非部分包含
步骤6: 点击"更新条件"保存该条件
步骤7: 根据需要添加更多条件
步骤8: 点击"更新规则"保存规则设置
步骤9: 点击"保存并关闭"完成规则配置

5. 配置提取列
--------------------------
步骤1: 点击"管理提取列"按钮打开提取列管理窗口
步骤2: 查看默认配置的列，或点击"添加列"创建新的提取列
步骤3: 设置列的显示名称和搜索名称
       - 显示名称：结果表格中显示的列标题
       - 搜索名称：用于在比对文件中查找对应列的关键词列表(多个值用逗号分隔)
步骤4: 配置列属性
       - 启用此列：是否在比对中使用此列
       - 作为主键列：标记为比对的关键列(如Part No)，结果中显示并用于匹配判断
步骤5: 点击"更新列"保存列配置
步骤6: 点击"保存并关闭"完成提取列配置

注意：必须有且仅有一个列被设置为主键列，用于比对结果的匹配判断。如果没有明确设置，系统会自动将第一个启用的列设为主键。

6. 执行比对操作
--------------------------
步骤1: 回到主界面，点击"开始比对"按钮
步骤2: 程序开始执行比对，状态栏会显示进度
步骤3: 比对完成后，结果将显示在表格中
步骤4: 弹出摘要对话框，显示各类结果的数量统计

7. 处理比对结果
--------------------------
• 查看结果表格，了解每一行的匹配情况
• 点击"导出结果"可将表格内容导出为Excel文件
• 浏览"匹配文件"文件夹查看自动生成的匹配文件
• 使用"合并匹配文件"功能整合匹配结果

四、高级功能详解
===============================

1. 自定义规则进阶用法
--------------------------
• 复杂条件设置：
  - 部分匹配：搜索值为字符串的一部分即可匹配
  - 精确匹配：要求完全相同，适合编码等精确字段
  - 正则表达式：用于复杂模式匹配，如：
    * ^ABC.*：以ABC开头的任意内容
    * .*XYZ$：以XYZ结尾的任意内容
    * [0-9]{3,5}：3到5位数字

• 多条件逻辑：
  - "满足所有条件"(AND逻辑)：要求同时满足所有设置的条件
  - "满足任一条件"(OR逻辑)：只需满足其中一个条件即可

• 规则组合策略：
  - 创建多个针对性规则，处理不同类型的数据
  - 根据优先级顺序排列规则
  - 系统按规则顺序检查，匹配第一个符合的规则

2. 提取列配置详解
--------------------------
提取列用于定义从比对文件中提取哪些数据，以及如何在结果中显示这些数据：

• 配置要点：
  - 显示名称：在结果表格中显示的列名称
  - 搜索名称：用于在比对文件中查找对应列的名称列表(支持多种可能的名称)
  - 主键列：用于比对结果判断的关键列，通常是Part No或类似唯一标识符
  - 启用/禁用：可临时禁用某些列，而不需要删除其配置

• 智能列匹配机制：
  - 精确匹配：先尝试完全匹配列名
  - 部分匹配：如果精确匹配失败，尝试部分匹配
  - 大小写不敏感：列名匹配忽略大小写
  - 自动处理空格：忽略列名首尾的空格差异

• 主键列的重要性：
  - 主键列(通常是Part No)是比对的核心
  - 系统使用主键列的值来判断是否匹配
  - 在结果表格中会显示为"对应文件+主键列名"
  - 每次比对必须有一个且仅有一个主键列

• 应用场景：
  - 跨文件比对：当不同文件使用不同列名表示相同数据时
  - 多语言支持：同时支持中英文等不同语言的列名
  - 灵活显示：自定义结果表格中的列名显示

3. 合并匹配文件功能
--------------------------
本功能提供两种合并方式：

• 合并为一张表：
  - 所有匹配文件数据合并到一个表中
  - 添加"来源文件"列标识数据来源
  - 适合需要统一分析处理的场景
  - 支持Excel和CSV两种输出格式

• 每个文件作为单独工作表：
  - 保留原始文件的数据结构
  - 每个文件内容单独放在一个工作表中
  - 工作表命名基于原文件名
  - 仅支持Excel格式输出
  - 适合需保留文件分类的场景

4. 特殊选项说明
--------------------------
• 无规则时提取所有行：
  - 默认情况下，当没有定义有效规则或规则没有条件时，系统返回"未找到符合条件的Part No"
  - 启用此选项后，系统将提取比对文件中的所有行用于比对
  - 适用于以下场景：
    * 简单数据比对，不需要复杂的筛选规则
    * 快速检查所有数据而非特定类型的数据
    * 希望先全量比对，再根据结果设计筛选规则

5. 结果分类说明
--------------------------
系统将比对结果分为四类：

• 匹配：找到的Part No与主文件完全相同（字符串完全相等）
• 不匹配：找到的所有Part No都与主文件中的不完全相同
• 其他结果：在有匹配结果的情况下，同时发现的其他不匹配Part No
• 错误：未找到对应文件或处理过程中出现异常

五、实际操作场景示例
===============================

1. 场景一：提取特定类别的物料记录
--------------------------
假设需要从多个物料清单中提取所有"电子元件"类别且为"进口"的物料：

步骤1: 设置规则名称为"电子元件-进口"
步骤2: 添加第一个条件 - 列名:"类别", 搜索值:"电子元件,电子器件"
步骤3: 添加第二个条件 - 列名:"产地", 搜索值:"进口,国外,海外"
步骤4: 选择"需要满足所有条件"选项
步骤5: 执行比对，系统将提取符合两个条件的记录

2. 场景二：识别符合多种规格的零件
--------------------------
如需提取符合不同规格要求的零件：

步骤1: 创建规则"A类规格零件"
步骤2: 添加条件 - 列名:"规格", 搜索值:"标准A,特殊A"
步骤3: 创建另一个规则"B类规格零件"
步骤4: 添加条件 - 列名:"规格", 搜索值:"标准B,特殊B"
步骤5: 执行比对，系统会分别识别两种类型的零件

3. 场景三：批量合并处理数据
--------------------------
收集并整合所有匹配记录：

步骤1: 完成数据比对
步骤2: 点击"合并匹配文件"
步骤3: 选择合并方式(单表或多工作表)
步骤4: 选择保存位置
步骤5: 在合并文件中进行后续数据分析处理

4. 场景四：全量数据比对
--------------------------
无需设置复杂规则，直接比对所有数据：

步骤1: 启用"无规则时提取所有行"选项
步骤2: 确保没有启用的比对规则，或创建一个没有条件的空规则
步骤3: 执行比对，系统会提取所有行进行对比
步骤4: 根据比对结果，进一步确定是否需要添加过滤规则

5. 场景五：应对不同文件格式的列名差异
--------------------------
当比对的文件来自不同来源，列名不一致时：

步骤1: 点击"管理提取列"
步骤2: 添加或编辑主键列(如Part No)
步骤3: 在搜索名称中输入所有可能的列名变体，如"part no,partnumber,料号,零件号"
步骤4: 保存设置并执行比对
步骤5: 系统会智能匹配各种不同名称的列，提取正确的数据

六、优化建议与常见问题
===============================

1. 提高比对效率的建议
--------------------------
• 合理组织比对文件夹，减少无关文件
• 设计精准的比对规则，避免过于宽泛的条件
• 对大文件先进行预处理或分割
• 优先使用精确匹配而非模糊匹配，提高速度
• 定期清理匹配文件夹，避免积累过多历史文件

2. 常见问题与解决方法
--------------------------
• 问题：找不到匹配项
  解决：检查比对规则是否正确配置，或启用"无规则时提取所有行"选项

• 问题：比对速度很慢
  解决：减少比对文件数量，精简比对规则，关闭不必要的应用程序

• 问题：CSV文件编码问题
  解决：程序会尝试多种编码，但如遇问题，建议先用Excel打开并另存为.xlsx格式

• 问题：某些文件无法读取
  解决：检查文件是否被其他程序占用，或存在格式损坏

• 问题：规则设置后不生效
  解决：确认规则已保存且已启用，检查条件设置是否正确

• 问题：无法找到正确的列
  解决：在"管理提取列"中为该列添加更多可能的搜索名称，增加匹配成功率

3. 数据准备建议
--------------------------
• 确保主文件中Model字段格式统一、无特殊字符
• 标准化各文件的表头命名，便于设置规则
• 对包含复杂格式的单元格，建议预处理为纯文本
• 备份原始数据，特别是首次使用本工具时

希望本指南能帮助您充分利用Excel比对工具的各项功能。如有更多问题，请联系技术支持。
"""
        
        # 插入文本
        help_text.insert(tk.END, help_content, "normal")
        
        # 添加样式
        # 查找并应用标题样式
        def apply_styles(tag, start_pattern, end_pattern="\n"):
            start_idx = "1.0"
            while True:
                start_idx = help_text.search(start_pattern, start_idx, tk.END)
                if not start_idx:
                    break
                end_idx = help_text.search(end_pattern, start_idx, tk.END)
                if not end_idx:
                    end_idx = help_text.index(tk.END)
                else:
                    end_idx = help_text.index(f"{end_idx}+{len(end_pattern)}c")
                help_text.tag_add(tag, start_idx, end_idx)
                start_idx = end_idx
        
        # 应用标题样式
        apply_styles("title", "Excel文件比对工具 - 详细使用说明")
        apply_styles("subtitle", "一、软件概述")
        apply_styles("subtitle", "二、工作原理与逻辑流程")
        apply_styles("subtitle", "三、详细操作流程")
        apply_styles("subtitle", "四、高级功能详解")
        apply_styles("subtitle", "五、实际操作场景示例")
        apply_styles("subtitle", "六、优化建议与常见问题")
        
        # 应用二级标题样式
        for i in range(1, 8):  # 更新数字以包含新添加的提取列配置部分
            apply_styles("heading", f"{i}. ", "\n")
        
        # 应用重点强调
        for term in ["核心比对逻辑", "规则引擎工作方式", "文件处理逻辑", "准备工作", "设置文件路径", 
                     "设置匹配选项", "配置比对规则", "配置提取列", "执行比对操作", "处理比对结果", 
                     "自定义规则进阶用法", "提取列配置详解", "合并匹配文件功能", "特殊选项说明", 
                     "结果分类说明", "场景一", "场景二", "场景三", "场景四", "场景五",
                     "提高比对效率的建议", "常见问题与解决方法", "数据准备建议"]:
            apply_styles("subheading", term, "\n")
        
        # 应用强调样式
        for emphasis in ["匹配", "不匹配", "其他结果", "错误", "部分匹配", "精确匹配", "正则表达式",
                         "满足所有条件", "满足任一条件", "合并为一张表", "每个文件作为单独工作表",
                         "无规则时提取所有行", "完全匹配模型名", "显示名称", "搜索名称", "主键列",
                         "启用/禁用", "精确匹配", "部分匹配"]:
            start_idx = "1.0"
            while True:
                start_idx = help_text.search(emphasis, start_idx, tk.END)
                if not start_idx:
                    break
                end_idx = f"{start_idx}+{len(emphasis)}c"
                # 避免标题中的样式被覆盖
                if not any(tag in help_text.tag_names(start_idx) for tag in ["title", "subtitle", "heading", "subheading"]):
                    help_text.tag_add("emphasis", start_idx, end_idx)
                start_idx = end_idx
        
        # 禁用编辑
        help_text.config(state=tk.DISABLED)
        
        # 添加关闭按钮
        ttk.Button(main_frame, text="关闭", command=help_dialog.destroy).pack(pady=10)
        
        # 窗口居中显示
        help_dialog.transient(self.root)
        help_dialog.update_idletasks()
        width = help_dialog.winfo_width()
        height = help_dialog.winfo_height()
        x = self.root.winfo_x() + (self.root.winfo_width() - width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - height) // 2
        help_dialog.geometry(f"{width}x{height}+{x}+{y}")
        
        # 设置为模态对话框
        help_dialog.grab_set()
        help_dialog.focus_set()

    def create_default_extract_columns(self):
        """创建默认的提取列配置"""
        # 创建一个Part No的提取列配置，设为主键
        part_no_column = ExtractColumn(
            name="Part No",
            search_names=["part no", "partno", "part number", "partnumber", "part_no", "part-no", "part", "零件号", "零件编号", "料号"],
            enabled=True,
            is_primary=True
        )
        self.extract_columns.append(part_no_column)
        
        # 添加一些常用的其他列配置示例
        self.extract_columns.append(ExtractColumn(
            name="描述",
            search_names=["description", "desc", "描述", "item desc", "物料描述"],
            enabled=True,
            is_primary=False
        ))
        
        self.extract_columns.append(ExtractColumn(
            name="规格",
            search_names=["specification", "spec", "规格", "item spec", "物料规格"],
            enabled=True,
            is_primary=False
        ))

    def manage_extract_columns(self):
        """打开提取列管理对话框"""
        # 创建一个新窗口
        columns_dialog = tk.Toplevel(self.root)
        columns_dialog.title("提取列管理")
        columns_dialog.minsize(750, 500)  # 设置最小尺寸
        columns_dialog.geometry("800x550")  # 设置初始尺寸
        
        # 创建主框架并使用padding
        main_frame = ttk.Frame(columns_dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 添加标题
        ttk.Label(main_frame, text="提取列管理", style='Header.TLabel').pack(pady=(0, 10))
        
        # 说明文本
        ttk.Label(main_frame, text="在此设置要从匹配文件中提取的列，系统将尝试根据以下配置搜索和匹配列名。", 
                  wraplength=700).pack(pady=(0, 10))
        
        # 创建水平分隔的左右面板 - 使用PanedWindow
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 左侧面板 - 列表框架
        list_frame = ttk.LabelFrame(paned, text="提取列列表", padding="5")
        
        # 右侧面板 - 详情框架
        detail_frame = ttk.LabelFrame(paned, text="列详情", padding="5")
        
        # 添加到PanedWindow
        paned.add(list_frame, weight=1)
        paned.add(detail_frame, weight=2)
        
        # 列表框架
        columns_frame = ttk.Frame(list_frame)
        columns_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 添加滚动条
        columns_scrollbar = ttk.Scrollbar(columns_frame)
        columns_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        columns_listbox = tk.Listbox(columns_frame, height=15, width=30, yscrollcommand=columns_scrollbar.set)
        columns_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        columns_scrollbar.config(command=columns_listbox.yview)
        
        # 列详情区域
        # 列属性框架
        column_props_frame = ttk.Frame(detail_frame)
        column_props_frame.pack(fill=tk.X, pady=5)
        
        # 标题和输入框使用网格布局使其对齐
        ttk.Label(column_props_frame, text="显示名称:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        column_name_var = tk.StringVar()
        ttk.Entry(column_props_frame, textvariable=column_name_var, width=30).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 搜索名称标签和文本框
        ttk.Label(column_props_frame, text="搜索名称:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        search_names_var = tk.StringVar()
        ttk.Entry(column_props_frame, textvariable=search_names_var, width=50).grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        
        # 说明文本
        ttk.Label(column_props_frame, text="多个搜索名称请用逗号分隔，系统将尝试匹配这些名称", 
                 wraplength=400).grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # 复选框
        enabled_var = tk.BooleanVar(value=True)
        is_primary_var = tk.BooleanVar(value=False)
        
        checkbox_frame = ttk.Frame(column_props_frame)
        checkbox_frame.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        ttk.Checkbutton(checkbox_frame, text="启用此列", variable=enabled_var).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Checkbutton(checkbox_frame, text="作为主键列（用于比对）", variable=is_primary_var).pack(side=tk.LEFT)
        
        # 说明文本
        ttk.Label(column_props_frame, text="注意: 必须有且仅有一个主键列用于比对结果", 
                 foreground="red").grid(row=4, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # 当前列索引变量
        current_column_index = {"value": -1}
        
        # 加载所有列配置到列表
        for column in self.extract_columns:
            display_text = column.name
            if column.is_primary:
                display_text += " (主键)"
            if not column.enabled:
                display_text += " (已禁用)"
            columns_listbox.insert(tk.END, display_text)
        
        # 加载列配置到编辑区域
        def load_column(index):
            if 0 <= index < len(self.extract_columns):
                column = self.extract_columns[index]
                current_column_index["value"] = index
                
                column_name_var.set(column.name)
                search_names_var.set(", ".join(column.search_names))
                enabled_var.set(column.enabled)
                is_primary_var.set(column.is_primary)
        
        def on_column_select(evt):
            """当选择列时"""
            selection = columns_listbox.curselection()
            if selection:
                index = selection[0]
                load_column(index)
        
        # 列操作函数
        def add_column():
            """添加新列配置"""
            new_column = ExtractColumn(
                name="新列",
                search_names=[""],
                enabled=True,
                is_primary=False
            )
            self.extract_columns.append(new_column)
            
            # 更新列表
            display_text = new_column.name
            columns_listbox.insert(tk.END, display_text)
            
            # 选择新列
            index = len(self.extract_columns) - 1
            columns_listbox.selection_clear(0, tk.END)
            columns_listbox.selection_set(index)
            columns_listbox.see(index)
            load_column(index)
        
        def update_column():
            """更新当前列配置"""
            index = current_column_index["value"]
            if 0 <= index < len(self.extract_columns):
                column = self.extract_columns[index]
                column.name = column_name_var.get()
                
                # 处理搜索名称，支持中英文逗号
                search_names_text = search_names_var.get()
                search_names_text = search_names_text.replace('，', ',')
                search_names = [v.strip() for v in search_names_text.split(',') if v.strip()]
                
                if search_names:
                    column.search_names = search_names
                
                column.enabled = enabled_var.get()
                was_primary = column.is_primary
                column.is_primary = is_primary_var.get()
                
                # 如果设置为主键，则其他列不能是主键
                if column.is_primary and not was_primary:
                    for i, other_column in enumerate(self.extract_columns):
                        if i != index and other_column.is_primary:
                            other_column.is_primary = False
                            # 更新列表显示
                            display_text = other_column.name
                            if not other_column.enabled:
                                display_text += " (已禁用)"
                            columns_listbox.delete(i)
                            columns_listbox.insert(i, display_text)
                
                # 更新列表显示
                display_text = column.name
                if column.is_primary:
                    display_text += " (主键)"
                if not column.enabled:
                    display_text += " (已禁用)"
                columns_listbox.delete(index)
                columns_listbox.insert(index, display_text)
                columns_listbox.selection_set(index)
        
        def delete_column():
            """删除当前列配置"""
            index = current_column_index["value"]
            if 0 <= index < len(self.extract_columns):
                # 检查是否删除了主键列
                was_primary = self.extract_columns[index].is_primary
                
                # 删除列配置
                del self.extract_columns[index]
                
                # 更新列表
                columns_listbox.delete(index)
                
                # 如果删除了主键列且还有其他列，则将第一个列设为主键
                if was_primary and self.extract_columns:
                    self.extract_columns[0].is_primary = True
                    # 更新显示
                    display_text = self.extract_columns[0].name + " (主键)"
                    if not self.extract_columns[0].enabled:
                        display_text += " (已禁用)"
                    columns_listbox.delete(0)
                    columns_listbox.insert(0, display_text)
                
                # 选择其他列
                if self.extract_columns:
                    new_index = min(index, len(self.extract_columns) - 1)
                    columns_listbox.selection_set(new_index)
                    load_column(new_index)
                else:
                    # 没有列了，清空编辑区
                    column_name_var.set("")
                    search_names_var.set("")
                    enabled_var.set(True)
                    is_primary_var.set(False)
                    current_column_index["value"] = -1
        
        # 绑定选择事件
        columns_listbox.bind('<<ListboxSelect>>', on_column_select)
        
        # 按钮框架
        button_frame = ttk.Frame(list_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_frame, text="添加列", command=add_column).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="更新列", command=update_column).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="删除列", command=delete_column).pack(side=tk.LEFT, padx=5)
        
        # 确认和取消按钮
        bottom_buttons_frame = ttk.Frame(main_frame)
        bottom_buttons_frame.pack(fill=tk.X, padx=10, pady=10)
        
        def on_save():
            # 确保至少有一个主键列
            has_primary = any(column.is_primary for column in self.extract_columns)
            if not has_primary and self.extract_columns:
                # 如果没有主键列，默认将第一个设为主键
                self.extract_columns[0].is_primary = True
                messagebox.showinfo("提示", "已自动将第一个列设置为主键列")
            
            # 如果有正在编辑的列，先保存它
            if current_column_index["value"] >= 0:
                update_column()
            self.save_settings()
            columns_dialog.destroy()
        
        ttk.Button(bottom_buttons_frame, text="保存并关闭", command=on_save).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_buttons_frame, text="取消", command=columns_dialog.destroy).pack(side=tk.RIGHT, padx=5)
        
        # 如果有列配置，默认选择第一个
        if self.extract_columns:
            columns_listbox.selection_set(0)
            load_column(0)
        
        # 使窗口在父窗口中居中
        columns_dialog.transient(self.root)
        columns_dialog.update_idletasks()
        width = columns_dialog.winfo_width()
        height = columns_dialog.winfo_height()
        x = self.root.winfo_x() + (self.root.winfo_width() - width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - height) // 2
        columns_dialog.geometry(f"{width}x{height}+{x}+{y}")
        
        # 设置为模态对话框
        columns_dialog.grab_set()
        columns_dialog.focus_set()
        columns_dialog.wait_window()

    def preview_file(self):
        """预览主文件的内容"""
        file_path = self.master_file_path.get()
        if not file_path:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
            
        # 检查文件是否存在
        if not os.path.exists(file_path):
            messagebox.showerror("错误", f"找不到文件: {file_path}")
            return
        
        try:
            # 读取文件
            if file_path.lower().endswith(('.xlsx', '.xls','.csv')):
                # Excel文件
                if self.master_sheet_name:
                    df = pd.read_excel(file_path, sheet_name=self.master_sheet_name)
                else:
                    df = pd.read_excel(file_path)
            else:
                # CSV文件
                try:
                    df = pd.read_csv(file_path, encoding='utf-8')
                except:
                    try:
                        df = pd.read_csv(file_path, encoding='gbk')
                    except:
                        df = pd.read_csv(file_path, encoding='gb18030')
        
            # 创建预览对话框
            preview_dialog = tk.Toplevel(self.root)
            preview_dialog.title(f"文件预览: {os.path.basename(file_path)}")
            preview_dialog.geometry("800x600")
            preview_dialog.transient(self.root)
            
            main_frame = ttk.Frame(preview_dialog, padding="10")
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # 显示前100行数据
            ttk.Label(main_frame, text=f"显示前 {min(100, len(df))} 行数据:").pack(anchor=tk.W, pady=(0, 5))
            
            # 创建表格展示数据
            preview_frame = ttk.Frame(main_frame)
            preview_frame.pack(fill=tk.BOTH, expand=True)
            
            # 添加滚动条
            v_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL)
            v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            h_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL)
            h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            
            # 创建Treeview
            cols = list(df.columns)
            preview_tree = ttk.Treeview(preview_frame, columns=cols, show="headings",
                                       yscrollcommand=v_scrollbar.set,
                                       xscrollcommand=h_scrollbar.set)
            
            v_scrollbar.config(command=preview_tree.yview)
            h_scrollbar.config(command=preview_tree.xview)
            
            # 设置列标题和宽度
            for col in cols:
                preview_tree.heading(col, text=str(col))
                # 设置适当的列宽
                max_width = len(str(col)) * 30
                for i in range(min(100, len(df))):
                    val = str(df.iloc[i][col])
                    width = len(val) * 10
                    max_width = max(max_width, min(width, 300))
                preview_tree.column(col, width=max_width)
            
            # 添加数据行
            for i in range(min(100, len(df))):
                values = [str(df.iloc[i][col]) for col in cols]
                preview_tree.insert("", tk.END, values=values)
            
            preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # 添加关闭按钮
            ttk.Button(main_frame, text="关闭", command=preview_dialog.destroy).pack(pady=10)
            
            # 窗口居中显示
            preview_dialog.update_idletasks()
            width = preview_dialog.winfo_width()
            height = preview_dialog.winfo_height()
            x = self.root.winfo_x() + (self.root.winfo_width() - width) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - height) // 2
            preview_dialog.geometry(f"{width}x{height}+{x}+{y}")
            
        except Exception as e:
            messagebox.showerror("错误", f"预览文件失败: {str(e)}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelComparator(root)
    
    # 在程序关闭时保存设置
    def on_closing():
        app.save_settings()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
