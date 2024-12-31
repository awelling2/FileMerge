import os
import time
from docxtpl import DocxTemplate
from config.config import (
    DOCX_OUTPUT_DIR,
    FILENAME_FORMATS,
    DEFAULT_OUTPUT_DIRS
)
import platform
from jinja2 import Template

class JsonToDocx:
    """JSON数据转Word文档处理类"""
    
    def __init__(self):
        self.last_generated_file = None
        
    def set_paths(self, template_path):
        """设置模板路径"""
        self.template_path = template_path
        
    def generate_docx_from_json(self, json_data):
        """从JSON数据生成Word文档"""
        try:
            # 添加设备参数
            device_params = {
                "设备参数": time.strftime("%Y年%m月%d日 %H时%M分%S秒") + " " + \
                         f"{platform.node()} ({platform.system()} {platform.release()})" + " " + \
                         os.path.basename(self.template_path)
            }
            json_data.update(device_params)
            
            # 添加模板名称用于文件名生成
            json_data['template_name'] = os.path.splitext(os.path.basename(self.template_path))[0]
            
            # 获取文件夹格式
            folder_format = json_data.get('folder_format')
            # 使用jinja2模板引擎处理文件夹格式
            folder_name = Template(folder_format).render(json_data)
            
            # 创建输出目录
            output_dir = os.path.join(DOCX_OUTPUT_DIR, folder_name)
            os.makedirs(output_dir, exist_ok=True)
            
            # 使用jinja2模板生成文件名
            format_template = Template(json_data.get('filename_format'))
            output_filename = format_template.render(**json_data)
            if not output_filename.endswith('.docx'):
                output_filename += '.docx'
            
            # 完整输出路径
            output_path = os.path.join(output_dir, output_filename)
            
            # 加载Word模板并渲染
            doc = DocxTemplate(self.template_path)
            doc.render(json_data)
            
            # 保存文件
            doc.save(output_path)
            self.last_generated_file = output_path
            return output_path
            
        except Exception as e:
            raise Exception(f"生成Word文档失败: {str(e)}") 