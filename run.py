#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FileMerge - 一个用于生成Word和Excel文档的模板合成工具

Copyright (C) 2024 Awelling2

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

import os
import sys
import ctypes

# 将项目根目录添加到 Python 路径
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)

from src.features.merge.merge_gui import MergeGUI
from config.config import SHOW_CONSOLE

def hide_console():
    """隐藏控制台窗口"""
    if os.name == 'nt':  # Windows
        kernel32 = ctypes.WinDLL('kernel32')
        user32 = ctypes.WinDLL('user32')
        hwnd = kernel32.GetConsoleWindow()
        if hwnd:
            user32.ShowWindow(hwnd, 0)
    elif sys.platform == 'darwin':  # macOS
        # macOS下不需要特殊处理，因为.app打包后默认不显示控制台
        pass
    else:  # Linux
        # Linux下可以通过启动参数来控制
        pass

def main():
    """程序入口"""
    try:
        # 根据配置决定是否显示控制台
        if not SHOW_CONSOLE:
            hide_console()
            
        # 创建必要的目录
        from config.config import (
            DOCUMENTS_DIR, TEMPLATE_DIR, EXCEL_TEMPLATE_DIR,
            TXT_DIR, FILE_MERGE_DIR, DOCX_OUTPUT_DIR, XLSX_OUTPUT_DIR
        )
        for directory in [DOCUMENTS_DIR, TEMPLATE_DIR, EXCEL_TEMPLATE_DIR,
                         TXT_DIR, FILE_MERGE_DIR, DOCX_OUTPUT_DIR, XLSX_OUTPUT_DIR]:
            os.makedirs(directory, exist_ok=True)
            
        # 启动GUI
        app = MergeGUI()
        app.win.mainloop()
        
    except Exception as e:
        import traceback
        print(f"程序启动失败: {str(e)}")
        print(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main() 