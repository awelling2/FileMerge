import json
import os
import platform
import subprocess
from .json_to_docx import JsonToDocx
from .json_to_excel import JsonToExcel

class FileHandler:
    """文件处理工具类"""
    def __init__(self):
        self.json_to_docx = JsonToDocx()
        self.json_to_excel = JsonToExcel()
        
    def read_json_data(self, file_path):
        """读取JSON配置文件"""
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            # 确保所有列表类型的数据都被正确处理
            for key, value in data.items():
                if isinstance(value, list):
                    # 确保列表不为空且第一个值有效
                    if value and value[0]:
                        data[key] = value
                    else:
                        data[key] = [""]  # 提供默认值
            return data
            
    def generate_docx(self, template_path, merge_data):
        """生成Word文档"""
        # 确保列表类型的数据只使用第一个值
        processed_data = {}
        for key, value in merge_data.items():
            if isinstance(value, list):
                processed_data[key] = value[0] if value else ""
            else:
                processed_data[key] = value
                
        self.json_to_docx.set_paths(template_path=template_path)
        return self.json_to_docx.generate_docx_from_json(processed_data)

    def generate_excel(self, template_path, merge_data):
        """生成Excel文档"""
        # 确保列表类型的数据只使用第一个值
        processed_data = {}
        for key, value in merge_data.items():
            if isinstance(value, list):
                processed_data[key] = value[0] if value else ""
            else:
                processed_data[key] = value
                
        self.json_to_excel.set_paths(template_path=template_path)
        return self.json_to_excel.generate_excel_from_json(processed_data)
        
    def open_file(self, filepath):
        """跨平台打开文件"""
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"文件不存在: {filepath}")
            
        try:
            if platform.system() == "Darwin":  # macOS
                subprocess.run(["open", filepath])
            elif platform.system() == "Windows":  # Windows
                os.startfile(filepath)
            else:  # Linux
                subprocess.run(["xdg-open", filepath])
        except Exception as e:
            raise Exception(f"无法打开文件: {str(e)}")

    def open_directory(self, directory):
        """跨平台打开目录"""
        if not os.path.exists(directory):
            raise FileNotFoundError(f"目录不存在: {directory}")
            
        try:
            if platform.system() == "Darwin":  # macOS
                subprocess.run(["open", directory])
            elif platform.system() == "Windows":  # Windows
                subprocess.run(["explorer", directory])
            else:  # Linux
                subprocess.run(["xdg-open", directory])
        except Exception as e:
            raise Exception(f"无法打开目录: {str(e)}")
        
    def get_latest_file(self, directory, extension=None):
        """获取目录中最新的文件"""
        if not os.path.exists(directory):
            raise FileNotFoundError(f"目录不存在: {directory}")
            
        files = []
        for file in os.listdir(directory):
            if extension and not file.endswith(extension):
                continue
            file_path = os.path.join(directory, file)
            if os.path.isfile(file_path):
                files.append(file_path)
            
        if not files:
            return None
            
        # 按修改时间排序，返回最新的文件
        return max(files, key=os.path.getmtime) 