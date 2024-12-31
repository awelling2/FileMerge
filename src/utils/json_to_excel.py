import os
from openpyxl import load_workbook
from config.config import (
    XLSX_OUTPUT_DIR,
    FILENAME_FORMATS,
    DEFAULT_OUTPUT_DIRS,
    FOLDER_FORMATS
)
from jinja2 import Template, Environment, BaseLoader

class JsonToExcel:
    """JSON数据转Excel文件处理类"""
    
    def __init__(self):
        self.last_generated_file = None
        # 创建jinja2环境
        self.env = Environment(loader=BaseLoader())
        
    def set_paths(self, template_path):
        """设置模板路径"""
        self.template_path = template_path
        
    def render_cell_content(self, template_str, data):
        """使用jinja2渲染单元格内容"""
        try:
            template = self.env.from_string(str(template_str))
            return template.render(**data)
        except Exception as e:
            # 如果渲染失败，返回原始内容
            return template_str
        
    def generate_excel_from_json(self, json_data):
        """根据JSON数据生成Excel文件"""
        try:
            # 获取文件夹格式
            folder_format = json_data.get('folder_format')
            # 使用jinja2模板引擎处理文件夹格式
            folder_name = Template(folder_format).render(json_data)
            
            # 创建输出目录
            output_dir = os.path.join(XLSX_OUTPUT_DIR, folder_name)
            os.makedirs(output_dir, exist_ok=True)
            
            # 获取文件名格式
            filename_format = json_data.get('filename_format')
            # 使用jinja2模板引擎处理文件名
            filename = Template(filename_format).render(json_data) + ".xlsx"
            
            # 生成完整的输出路径
            output_path = os.path.join(output_dir, filename)
            
            # 加载Excel模板
            wb = load_workbook(self.template_path)
            ws = wb.active
            
            # 遍历所有单元格，使用jinja2渲染内容
            for row in ws.rows:
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        template = Template(cell.value)
                        cell.value = template.render(**json_data)
            
            # 保存文件
            wb.save(output_path)
            # 保存最后生成的文件路径
            self.last_generated_file = output_path
            return output_path
            
        except Exception as e:
            raise Exception(f"生成Excel文档失败: {str(e)}") 