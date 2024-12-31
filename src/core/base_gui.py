import tkinter as tk
from ttkbootstrap import ttk
from ttkbootstrap.style import Style
import time
import os
import platform
import subprocess
import logging
from PIL import Image, ImageTk
from config.config import (
    GUI_TITLE, GUI_WIDTH, GUI_HEIGHT,
    PROJECT_ROOT, DEFAULT_THEME
)

class BaseGUI:
    """基础GUI类，包含共用的GUI初始化和通用方法"""
    def __init__(self):
        self.style = Style(theme=DEFAULT_THEME)
        self.win = self.style.master
        self.setup_root()
        self._create_menu()
        self._create_time_label()
        
    def setup_root(self):
        """初始化根窗口"""
        self.win.title(GUI_TITLE)
        self.win.attributes("-alpha", 1, "-topmost", 1)
        self.win.geometry(f"{GUI_WIDTH}x{GUI_HEIGHT}")
        
        # 获取屏幕尺寸
        screen_width = self.win.winfo_screenwidth()
        screen_height = self.win.winfo_screenheight()
        
        # 计算窗口位置 - 右上角
        x = screen_width - GUI_WIDTH
        y = 0
        
        # 设置窗口位置
        self.win.geometry(f"{GUI_WIDTH}x{GUI_HEIGHT}+{x}+{y}")
        
        # 设置软件图标
        try:
            icon_path = os.path.join(PROJECT_ROOT, "photos", "icon.ico")
            if os.path.exists(icon_path):
                self.win.iconbitmap(icon_path)
        except Exception as e:
            logging.error(f"加载图标失败: {str(e)}")
        
    def _create_menu(self):
        self.menu_bar = tk.Menu(self.win)
        self.menu_bar.add_cascade(label="打开合成病例目录", command=self.open_case_directory)
        self.win.config(menu=self.menu_bar)
        
    def _create_time_label(self):
        self.time_label = tk.Label(self.win, text=time.strftime('现在系统时间是：%Y-%m-%d %H:%M:%S'))
        # self.time_label.grid(row=0, column=0, sticky="w")
        self.refresh_time()
        
    def refresh_time(self):
        self.time_label.config(text=time.strftime('现在系统时间是：%Y-%m-%d %H:%M:%S'))
        self.win.after(1000, self.refresh_time)
        
    def open_directory(self, directory):
        """通用打开目录方法"""
        if not os.path.exists(directory):
            tk.messagebox.showerror("错误", f"目录 {directory} 不存在")
            return
            
        if platform.system() == "Darwin":  # macOS
            subprocess.run(["open", directory])
        elif platform.system() == "Windows":  # Windows
            subprocess.run(["explorer", directory])
        else:  # Linux
            subprocess.run(["xdg-open", directory])
            
    def open_case_directory(self):
        """需在子类中实现"""
        pass 