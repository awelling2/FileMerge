import tkinter as tk
from tkinter import filedialog
from ttkbootstrap import ttk
from ttkbootstrap.style import Style
import os
import sys
from tkinter import messagebox
from config.config import (
    GUI_TITLE, GUI_WIDTH, GUI_HEIGHT,
    TEMPLATE_DIR, EXCEL_TEMPLATE_DIR, TXT_DIR, FILE_MERGE_DIR,
    DOCX_TEMPLATE_TXT, DEFAULT_DOCX_TEMPLATE, DEFAULT_EXCEL_TEMPLATE,
    WECHAT_PAY_IMAGE, REQUIRED_FIELDS, ABOUT_TEXT,
    PROJECT_ROOT, THEMES, DEFAULT_THEME, MERGE_MODES, DEFAULT_MERGE_MODE,
    FILENAME_FORMATS, FOLDER_FORMATS
)
import time
import platform
from PIL import Image, ImageTk

from src.core.base_gui import BaseGUI
from src.utils.file_handler import FileHandler

# 获取项目根目录路径
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../.."))

class MergeGUI(BaseGUI):
    def __init__(self):
        """初始化GUI"""
        # 初始化基本变量
        self.current_theme = DEFAULT_THEME
        self.current_merge_mode = DEFAULT_MERGE_MODE
        self.donation_window = None
        
        # 初始化父类
        super().__init__()
        
        self.file_handler = FileHandler()
        
        # 初始化变量
        self.case_keyword = {}
        self.widget_vars = {}
        self.current_config_path = MERGE_MODES[DEFAULT_MERGE_MODE]["config"]  # 使用对应的配置文件
        self.current_template_path = MERGE_MODES[DEFAULT_MERGE_MODE]["template"]
        
        # 设置UI
        self._create_menu()
        self._create_time_frame()
        self._create_merge_mode_frame()
        self._create_file_selection()
        self._create_data_labelframe()
        self._create_action_buttons()
        
        # 绑定快捷键
        self._bind_shortcuts()
        
        # 设置窗口大小调整
        self.win.resizable(True, True)
        
        # 自动加载默认配置
        try:
            self.case_keyword = self.file_handler.read_json_data(self.current_config_path)
            self._create_data_widgets()
        except Exception as e:
            messagebox.showerror("错误", f"读取默认配置失败: {str(e)}")

    def _create_time_frame(self):
        """创建时间示框架"""
        time_frame = ttk.LabelFrame(self.win, text="系统时间")
        time_frame.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=5)
        
        self.time_label = tk.Label(time_frame, text=time.strftime('%Y-%m-%d %H:%M:%S'))
        self.time_label.pack(side="left", padx=10, pady=5)
        self.refresh_time()

    def _create_merge_mode_frame(self):
        """创建默认模版选择框架"""
        mode_frame = ttk.LabelFrame(self.win, text="默认模版")
        mode_frame.grid(row=0, column=0, padx=280, pady=5)
        
        self.merge_mode_var = tk.StringVar(value=MERGE_MODES[DEFAULT_MERGE_MODE]["name"])
        
        # 创建单选按钮，竖向排列
        for mode, config in MERGE_MODES.items():
            ttk.Radiobutton(
                mode_frame,
                text=config["name"],
                variable=self.merge_mode_var,
                value=config["name"],
                command=lambda m=mode: self.change_merge_mode(m)
            ).pack(side="top", padx=5, pady=2)  # 为竖向排列，添加垂直间距

    def change_merge_mode(self, mode):
        """切换合成模式"""
        try:
            self.current_merge_mode = mode
            # 更新当前配置文件和模板文件路径
            self.current_config_path = MERGE_MODES[mode]["config"]
            self.current_template_path = MERGE_MODES[mode]["template"]
            
            # 更新显示
            self.file_path_var.set(os.path.basename(self.current_config_path))
            self.template_var.set(os.path.basename(self.current_template_path))
            
            # 加载新的配置数据
            self.case_keyword = self.file_handler.read_json_data(self.current_config_path)
            
            # 清空并重新创建数据输入区域
            for widget in self.labelframe_data.winfo_children():
                widget.destroy()
            self._create_data_widgets()
            
            # 更新文件名格式变量
            if mode == "docx":
                self.docx_filename_format_var.set(list(FILENAME_FORMATS["docx"].keys())[0])
                self.docx_folder_format_var.set(list(FOLDER_FORMATS["docx"].keys())[0])
            else:  # excel
                self.excel_filename_format_var.set(list(FILENAME_FORMATS["excel"].keys())[0])
                self.excel_folder_format_var.set(list(FOLDER_FORMATS["excel"].keys())[0])
            
        except Exception as e:
            messagebox.showerror("错误", f"切换模式失败: {str(e)}")
            # 发生错误时恢复到默认模式
            self.merge_mode_var.set(MERGE_MODES[DEFAULT_MERGE_MODE]["name"])

    def _create_file_selection(self):
        """创建文件选择框架"""
        config_frame = ttk.LabelFrame(self.win, text="文件选择")
        config_frame.grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=1)
        
        # txt配置文件选择
        file_label = tk.Label(config_frame, text="配置：")
        file_label.grid(row=0, column=0, sticky="w", padx=10)

        self.file_path_var = tk.StringVar(value=os.path.basename(MERGE_MODES[DEFAULT_MERGE_MODE]["config"]))
        file_entry = ttk.Entry(config_frame, textvariable=self.file_path_var, width=20, state='readonly')
        file_entry.grid(row=0, column=1, sticky="w", padx=5)

        self.browse_config_btn = ttk.Button(config_frame, text="浏览", command=self.open_file)
        self.browse_config_btn.grid(row=0, column=2, padx=5)

        self.open_config_btn = ttk.Button(config_frame, text="打开", command=self.open_current_config)
        self.open_config_btn.grid(row=0, column=3, padx=5)

        # 模板文件选择
        template_label = tk.Label(config_frame, text="模板：")
        template_label.grid(row=1, column=0, sticky="w", padx=10)

        self.template_var = tk.StringVar(value=os.path.basename(self.current_template_path))
        template_entry = ttk.Entry(config_frame, textvariable=self.template_var, width=20, state='readonly')
        template_entry.grid(row=1, column=1, sticky="w", padx=5)

        self.browse_template_btn = ttk.Button(config_frame, text="浏览", command=self.open_template)
        self.browse_template_btn.grid(row=1, column=2, padx=5)

        self.open_template_btn = ttk.Button(config_frame, text="打开", command=self.open_current_template)
        self.open_template_btn.grid(row=1, column=3, padx=5)

    def _create_data_labelframe(self):
        """创建数据输入框架"""
        self.labelframe_data = tk.LabelFrame(self.win, text="数据输入")
        self.labelframe_data.grid(row=4, column=0, sticky="w", padx=10)
        self._create_data_widgets()

    def _create_data_widgets(self):
        # 清除现有的小部件
        for widget in self.labelframe_data.winfo_children():
            widget.destroy()
            
        # 创建新的小部件
        for i, key in enumerate(self.case_keyword):
            tk.Label(self.labelframe_data, text=f"{key}：").grid(row=i, column=0, sticky="e")
            self._create_input_widget(key, i)
            
    def _create_input_widget(self, key, row):
        var = tk.StringVar()
        # 设置初始值
        if isinstance(self.case_keyword[key], list):
            var.set(self.case_keyword[key][0])  # 列表的第一个值作为默认值
        else:
            var.set(self.case_keyword[key])  # 直接使用值
        self.widget_vars[key] = var
        
        if isinstance(self.case_keyword[key], list):
            # 如果是列表类型，创建下拉框
            widget = ttk.Combobox(
                self.labelframe_data,
                textvariable=var,
                values=self.case_keyword[key],
                width=27
            )
            # 绑定选择事件
            widget.bind('<<ComboboxSelected>>', lambda e, k=key: self._on_value_change(k))
            widget.bind('<FocusOut>', lambda e, k=key: self._on_value_change(k))
            widget.bind('<Return>', lambda e, k=key: self._on_value_change(k))
        else:
            # 如果不是列表，创建普通输入框
            widget = tk.Entry(
                self.labelframe_data,
                textvariable=var,
                width=30
            )
            # 绑定输入事件
            widget.bind('<FocusOut>', lambda e, k=key: self._on_value_change(k))
            widget.bind('<Return>', lambda e, k=key: self._on_value_change(k))
            
        widget.grid(row=row, column=1, sticky="w", padx=10)

    def _on_value_change(self, key):
        """当输入值改变时更新数据"""
        value = self.widget_vars[key].get()
        if isinstance(self.case_keyword[key], list):
            # 如果是列表类型，保持列表结构，但将当前选择的值放在第一
            if value in self.case_keyword[key]:
                values = self.case_keyword[key].copy()
                values.remove(value)
                values.insert(0, value)
                self.case_keyword[key] = values
        else:
            # 如果是普通值，直接更新
            self.case_keyword[key] = value

    def _create_action_buttons(self):
        """创建操作按钮框架"""
        button_frame = ttk.LabelFrame(self.win, text="操作")
        button_frame.grid(row=5, column=0, sticky="w", padx=10, pady=1)
        
        merge_btn = ttk.Button(button_frame, text="合成(Ctrl+M)", command=self.merge_file)
        merge_btn.grid(row=0, column=0, padx=10, pady=1)
        
        open_btn = ttk.Button(button_frame, text="打开文件(Ctrl+O)", command=self.open_merged_file)
        open_btn.grid(row=0, column=1, padx=10, pady=1)

    def open_file(self):
        """打开配置文件，固定在 txt 目录"""
        path = filedialog.askopenfilename(
            initialdir=TXT_DIR,
            title="选择配置文件",
            filetypes=[("Text files", "*.txt")]
        )
        if not path:
            return
            
        self.current_config_path = path
        filename = os.path.basename(path)
        self.file_path_var.set(filename)
        
        try:
            self.case_keyword = self.file_handler.read_json_data(path)
            self.labelframe_data.config(text="关键词数据")  # 使用固定标题
            self._create_data_widgets()
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")
            
    def open_template(self):
        """打开模板文件选择对话框"""
        path = filedialog.askopenfilename(
            initialdir=TEMPLATE_DIR if self.current_merge_mode == "docx" else EXCEL_TEMPLATE_DIR,
            title="选择模板文件",
            filetypes=MERGE_MODES[self.current_merge_mode]["filetypes"]
        )
        if not path:
            return
            
        self.current_template_path = path
        self.template_var.set(os.path.basename(path))
        
        # 根据选择的文件类型自动切换模式
        new_mode = None
        if path.lower().endswith('.docx'):
            new_mode = "docx"
            self.merge_mode_var.set(MERGE_MODES["docx"]["name"])
            # 切换到Word配置文件
            self.current_config_path = MERGE_MODES["docx"]["config"]
            self.file_path_var.set(os.path.basename(self.current_config_path))
        elif path.lower().endswith('.xlsx'):
            new_mode = "excel"
            self.merge_mode_var.set(MERGE_MODES["excel"]["name"])
            # 切换到Excel配置文件
            self.current_config_path = MERGE_MODES["excel"]["config"]
            self.file_path_var.set(os.path.basename(self.current_config_path))
            
        if new_mode and new_mode != self.current_merge_mode:
            self.current_merge_mode = new_mode
            # 更新文件名格式菜单
            default_format = list(FILENAME_FORMATS[new_mode].keys())[0]
            self.filename_format_var.set(default_format)
            
        # 重新加载配置数据
        try:
            self.case_keyword = self.file_handler.read_json_data(self.current_config_path)
            self._create_data_widgets()
        except Exception as e:
            messagebox.showerror("错误", f"读取配置文件失败: {str(e)}")

    def open_current_template(self):
        """打开当前模板文件"""
        if not self.current_template_path:
            messagebox.showerror("错误", "请先选择模板文件")
            return
            
        try:
            self.file_handler.open_file(self.current_template_path)
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def merge_file(self):
        """合成文件"""
        try:
            if not self.current_template_path or not os.path.exists(self.current_template_path):
                messagebox.showerror("错误", "请先选择模板文件")
                return
                
            if not self._validate_data():
                return
                
            merge_data = self.update_data_from_gui()
            
            if self.current_merge_mode == "docx":
                # 生成Word文档
                output_file = self.file_handler.generate_docx(
                    self.current_template_path,
                    merge_data
                )
                messagebox.showinfo("提示", f"病例合成完成，文件保存路径：\n{output_file}")
            else:  # excel
                # 生成Excel文档
                output_file = self.file_handler.generate_excel(
                    self.current_template_path,
                    merge_data
                )
                messagebox.showinfo("提示", f"Excel文件生成完成，保存路径：\n{output_file}")
                
        except Exception as e:
            messagebox.showerror("错误", f"合成失败: {str(e)}")

    def update_data_from_gui(self):
        """从GUI获取更新后的数据"""
        updated_data = {}
        
        # 获取所有输入框的值
        for key in self.case_keyword:
            if key not in self.widget_vars:
                continue
                
            widget = self.widget_vars[key]
            value = widget.get()
            
            # 处理列表类型的数据
            if isinstance(self.case_keyword[key], list):
                if value in self.case_keyword[key]:
                    # 将选中的值放在列表第一位
                    values = self.case_keyword[key].copy()
                    values.remove(value)
                    values.insert(0, value)
                    updated_data[key] = value  # 只保存选中的值，不保存整个列表
                else:
                    # 如果值不在列表中，使用第一个值
                    updated_data[key] = self.case_keyword[key][0] if self.case_keyword[key] else ""
            else:
                # 如果是必填字段且为空，抛出异常
                if key in REQUIRED_FIELDS and not value.strip():
                    raise ValueError(f"请填写必填字段：{key}")
                updated_data[key] = value.strip() if isinstance(value, str) else value
        
        # 添加文件名格式
        if self.current_merge_mode == "docx":
            format_name = self.docx_filename_format_var.get()
            updated_data['filename_format'] = FILENAME_FORMATS["docx"][format_name]
            # 添加文件夹格式
            folder_format_name = self.docx_folder_format_var.get()
            updated_data['folder_format'] = FOLDER_FORMATS["docx"][folder_format_name]
        else:  # excel
            format_name = self.excel_filename_format_var.get()
            updated_data['filename_format'] = FILENAME_FORMATS["excel"][format_name]
            # 添加文件夹格式
            folder_format_name = self.excel_folder_format_var.get()
            updated_data['folder_format'] = FOLDER_FORMATS["excel"][folder_format_name]
        
        return updated_data
        
    def open_merged_file(self):
        """打开最新合成的文件"""
        try:
            if self.current_merge_mode == "docx":
                last_file = self.file_handler.json_to_docx.last_generated_file
            else:  # excel
                last_file = self.file_handler.json_to_excel.last_generated_file
                
            if last_file and os.path.exists(last_file):
                self.file_handler.open_file(last_file)
            else:
                messagebox.showerror("错误", "未找到最近合成的文件")
        except Exception as e:
            messagebox.showerror("错误", str(e))
            
    def open_current_config(self):
        """打开当前配置文件"""
        if not self.current_config_path:
            messagebox.showerror("错误", "请先选择配置文件")
            return
            
        try:
            self.file_handler.open_file(self.current_config_path)
        except Exception as e:
            messagebox.showerror("错误", str(e))
            
    def open_case_directory(self):
        """打开输出目录"""
        try:
            self.file_handler.open_directory(FILE_MERGE_DIR)
        except Exception as e:
            messagebox.showerror("错误", str(e))
            
    def toggle_default_config(self):
        """切换默认配置"""
        if self.use_default_var.get():
            # 使用默认配置
            self.current_config_path = MERGE_MODES[DEFAULT_MERGE_MODE]["config"]
            self.current_template_path = MERGE_MODES[DEFAULT_MERGE_MODE]["template"]
            
            # 更新显示
            self.file_path_var.set(os.path.basename(self.current_config_path))
            self.template_var.set(os.path.basename(self.current_template_path))
            
            # 禁用浏览按钮
            self.browse_config_btn.configure(state='disabled')
            self.browse_template_btn.configure(state='disabled')
            
            # 加载默认配��
            try:
                self.case_keyword = self.file_handler.read_json_data(self.current_config_path)
                self._create_data_widgets()
            except Exception as e:
                messagebox.showerror("错误", f"读取默认配置失败: {str(e)}")
                # 错误时取消选中复选框
                self.use_default_var.set(False)
                self.browse_config_btn.configure(state='normal')
                self.browse_template_btn.configure(state='normal')
        else:
            # 启用浏览按钮
            self.browse_config_btn.configure(state='normal')
            self.browse_template_btn.configure(state='normal')
            
            # 清空当前选择
            self.current_config_path = ""
            self.current_template_path = ""
            self.file_path_var.set("")
            self.template_var.set("")
            
            # 清空数据输入区域
            for widget in self.labelframe_data.winfo_children():
                widget.destroy()

    def _create_menu(self):
        """创建菜单栏"""
        self.menu_bar = tk.Menu(self.win)
        
        # 文件菜单
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        file_menu.add_command(label="打开输出目录", command=self.open_case_directory)
        file_menu.add_separator()
        file_menu.add_command(label="打开配置模板目录", command=lambda: self.file_handler.open_directory(TXT_DIR))
        file_menu.add_command(label="打开Word模板目录", command=lambda: self.open_template_directory("docx"))
        file_menu.add_command(label="打开Excel模板目录", command=lambda: self.open_template_directory("excel"))
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.win.quit)
        
        # 设置菜单
        settings_menu = tk.Menu(self.menu_bar, tearoff=0)
        
        # Word设置子菜单
        docx_menu = tk.Menu(settings_menu, tearoff=0)
        
        # Word文件名格式子菜单
        docx_filename_menu = tk.Menu(docx_menu, tearoff=0)
        self.docx_filename_format_var = tk.StringVar(value=list(FILENAME_FORMATS["docx"].keys())[0])
        for format_name in FILENAME_FORMATS["docx"].keys():
            docx_filename_menu.add_radiobutton(
                label=format_name,
                variable=self.docx_filename_format_var,
                value=format_name
            )
        docx_menu.add_cascade(label="文件名格式(仅支持docx)", menu=docx_filename_menu)
        
        # Word文件夹格式子菜单
        docx_folder_menu = tk.Menu(docx_menu, tearoff=0)
        self.docx_folder_format_var = tk.StringVar(value=list(FOLDER_FORMATS["docx"].keys())[0])
        for format_name in FOLDER_FORMATS["docx"].keys():
            docx_folder_menu.add_radiobutton(
                label=format_name,
                variable=self.docx_folder_format_var,
                value=format_name
            )
        docx_menu.add_cascade(label="文件夹格式", menu=docx_folder_menu)
        
        settings_menu.add_cascade(label="Word设置", menu=docx_menu)
        
        # Excel设置子菜单
        excel_menu = tk.Menu(settings_menu, tearoff=0)
        
        # Excel文件名格式子菜单
        excel_filename_menu = tk.Menu(excel_menu, tearoff=0)
        self.excel_filename_format_var = tk.StringVar(value=list(FILENAME_FORMATS["excel"].keys())[0])
        for format_name in FILENAME_FORMATS["excel"].keys():
            excel_filename_menu.add_radiobutton(
                label=format_name,
                variable=self.excel_filename_format_var,
                value=format_name
            )
        excel_menu.add_cascade(label="文件名格式", menu=excel_filename_menu)
        
        # Excel文件夹格式子菜单
        excel_folder_menu = tk.Menu(excel_menu, tearoff=0)
        self.excel_folder_format_var = tk.StringVar(value=list(FOLDER_FORMATS["excel"].keys())[0])
        for format_name in FOLDER_FORMATS["excel"].keys():
            excel_folder_menu.add_radiobutton(
                label=format_name,
                variable=self.excel_folder_format_var,
                value=format_name
            )
        excel_menu.add_cascade(label="文件夹格式", menu=excel_folder_menu)
        
        settings_menu.add_cascade(label="Excel设置", menu=excel_menu)
        
        # 帮助菜单
        help_menu = tk.Menu(self.menu_bar, tearoff=0)
        help_menu.add_command(label="快捷键说明", command=self.show_shortcuts_help)
        
        self.menu_bar.add_cascade(label="文件", menu=file_menu)
        self.menu_bar.add_cascade(label="设置", menu=settings_menu)
        self.menu_bar.add_cascade(label="帮助", menu=help_menu)
        self.menu_bar.add_command(label="关于", command=self.show_about)
        self.menu_bar.add_command(label="捐助", command=self.show_donation)
        self.win.config(menu=self.menu_bar)
        
    def open_template_directory(self, mode):
        """打开模板目录"""
        try:
            if mode == "docx":
                self.file_handler.open_directory(TEMPLATE_DIR)
            else:  # excel
                self.file_handler.open_directory(EXCEL_TEMPLATE_DIR)
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def change_theme(self, theme_name):
        """切换主题"""
        try:
            # 保存当前窗口位置
            current_x = self.win.winfo_x()
            current_y = self.win.winfo_y()
            
            # 创建新的Style对象
            self.style = Style(theme=theme_name)
            # 更新窗口引用
            self.win = self.style.master
            
            # 先创建所有UI元素
            self._create_menu()
            self._create_time_frame()
            self._create_merge_mode_frame()
            self._create_file_selection()
            self._create_data_labelframe()
            self._create_action_buttons()
            
            # 重新绑定快捷键
            self._bind_shortcuts()
            
            # 最后设置窗口属性
            self.win.title(GUI_TITLE)
            self.win.attributes("-alpha", 1, "-topmost", 1)
            # 确保窗口大小正确
            self.win.geometry(f"{GUI_WIDTH}x{GUI_HEIGHT}+{current_x}+{current_y}")
            self.win.update_idletasks()  # 强制更新窗口
            
            # 更新当前主题
            self.current_theme = theme_name
            messagebox.showinfo("提示", f"主题已切换为：{theme_name}")
        except Exception as e:
            messagebox.showerror("错误", f"切换主题失败：{str(e)}")

    def show_shortcuts_help(self):
        """显示快捷键说明"""
        shortcut_text = "Command" if platform.system() == "Darwin" else "Ctrl"
        help_text = SHORTCUTS_HELP.format(shortcut=shortcut_text)
        messagebox.showinfo("使用帮助", help_text)

    def _bind_shortcuts(self):
        """绑定快捷键"""
        # 判断操作系统
        if platform.system() == "Darwin":  # macOS
            self.win.bind('<Command-m>', lambda e: self.merge_file())
            self.win.bind('<Command-o>', lambda e: self.open_merged_file())
        else:  # Windows/Linux
            self.win.bind('<Control-m>', lambda e: self.merge_file())
            self.win.bind('<Control-o>', lambda e: self.open_merged_file())

    def refresh_time(self):
        """更新时间显示"""
        self.time_label.config(text=time.strftime('现在系统时间是：%Y-%m-%d %H:%M:%S'))
        self.win.after(1000, self.refresh_time) 

    def _validate_data(self):
        """验证数据"""
        try:
            # 检查必填字段
            for key, widget in self.widget_vars.items():
                if isinstance(widget, (ttk.Entry, ttk.Combobox)):
                    value = widget.get().strip()
                    if key in REQUIRED_FIELDS and not value:
                        messagebox.showerror("错误", f"请填写必填字段：{key}")
                        return False
            return True
        except Exception as e:
            messagebox.showerror("错误", f"数据验证失败: {str(e)}")
            return False

    def show_about(self):
        """显示关于信息"""
        messagebox.showinfo("关于", ABOUT_TEXT)

    def show_donation(self):
        """显示捐赠信息"""
        if self.donation_window is not None:
            self.donation_window.lift()
            return
            
        try:
            self.donation_window = tk.Toplevel(self.win)
            self.donation_window.title("微信捐助")
            self.donation_window.attributes('-topmost', True)
            self.donation_window.protocol("WM_DELETE_WINDOW", self._on_donation_window_close)
            
            photo_path = WECHAT_PAY_IMAGE  # 使用配置中的图片路径
            if os.path.exists(photo_path):
                img = Image.open(photo_path)
                img = img.resize((300, 400), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                label = tk.Label(self.donation_window, image=photo)
                label.image = photo
                label.pack()
            else:
                tk.Label(self.donation_window, text="图片加载失败").pack()
                
        except Exception as e:
            messagebox.showerror("错误", f"无法加载捐助图片: {str(e)}")
            self.donation_window = None

    def _on_donation_window_close(self):
        """捐助窗口关闭时的处理"""
        self.donation_window.destroy()
        self.donation_window = None 

    def merge_excel(self):
        """合成Excel文件"""
        if not self._validate_data():
            return
            
        merge_data = self.update_data_from_gui()
        
        try:
            output_file = self.file_handler.generate_excel(
                DEFAULT_EXCEL_TEMPLATE,
                merge_data
            )
            messagebox.showinfo("提示", f"Excel合成完成，文件保存路径：\n{output_file}")
        except Exception as e:
            messagebox.showerror("错误", f"合成失败: {str(e)}")
            
    def open_merged_excel(self):
        """打开最新合成的Excel文件"""
        try:
            if hasattr(self.file_handler.json_to_excel, 'last_generated_file') and \
               self.file_handler.json_to_excel.last_generated_file:
                self.file_handler.open_file(self.file_handler.json_to_excel.last_generated_file)
            else:
                messagebox.showerror("错误", "未找到最近合成的Excel文件")
        except Exception as e:
            messagebox.showerror("错误", str(e))
