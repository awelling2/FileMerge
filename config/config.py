import os

# 获取项目根目录
PROJECT_ROOT = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))

# 版本信息
VERSION = "1.0.2"
AUTHOR = "Awelling2"
EMAIL = "zx824548472@gmail.com"

# GUI配置
GUI_TITLE = "FileMerge " + VERSION
GUI_WIDTH = 400
GUI_HEIGHT = 840

# 目录配置
DOCUMENTS_DIR = os.path.join(PROJECT_ROOT, "documents")
TEMPLATE_DIR = os.path.join(DOCUMENTS_DIR, "template_docx")
EXCEL_TEMPLATE_DIR = os.path.join(DOCUMENTS_DIR, "template_excel")
TXT_DIR = os.path.join(DOCUMENTS_DIR, "txt")
FILE_MERGE_DIR = os.path.join(DOCUMENTS_DIR, "file_merge")  # 新的主输出目录
DOCX_OUTPUT_DIR = os.path.join(FILE_MERGE_DIR, "docx")  # Word文档输出目录
XLSX_OUTPUT_DIR = os.path.join(FILE_MERGE_DIR, "xlsx")  # Excel文档输出目录
PHOTOS_DIR = os.path.join(PROJECT_ROOT, "photos")

# 默认文件
DOCX_TEMPLATE_TXT = os.path.join(TXT_DIR, "docx_template.txt")  # Word模板配置文件
XLSX_TEMPLATE_TXT = os.path.join(TXT_DIR, "xlsx_template.txt")  # Excel模板配置文件
DEFAULT_DOCX_TEMPLATE = os.path.join(TEMPLATE_DIR, "template.docx")
DEFAULT_EXCEL_TEMPLATE = os.path.join(EXCEL_TEMPLATE_DIR, "template.xlsx")

# 图片资源
WECHAT_PAY_IMAGE = os.path.join(PHOTOS_DIR, "wechat_pay.jpg")
ICON_IMAGE = os.path.join(PHOTOS_DIR, "icon.ico")

# 必填字段
REQUIRED_FIELDS = ['序号', '姓名', '手术眼别', '入院日期', '手术日期', '出院日期']

# 主题配置
THEMES = [
    "journal",    # 默认主题
    "flatly",     # 扁平化主题
    "darkly",     # 深色主题
    "solar",      # 太阳主题
    "superhero",  # 超级英雄主题
    "cyborg",     # 机器人主题
    "litera",     # 文学主题
    "minty",      # 薄荷主题
    "lumen",      # 光明主题
    "pulse"       # 脉冲主题
]
DEFAULT_THEME = "journal"

# 控制台设置
SHOW_CONSOLE = False
# 合成模式
MERGE_MODES = {
    "docx": {
        "name": "Word模板",
        "template": os.path.join(TEMPLATE_DIR, "template.docx"),
        "filetypes": [("Word documents", "*.docx")],
        "config": DOCX_TEMPLATE_TXT  # 添加对应的配置文件
    },
    "excel": {
        "name": "Excel模板",
        "template": os.path.join(EXCEL_TEMPLATE_DIR, "template.xlsx"),
        "filetypes": [("Excel files", "*.xlsx")],
        "config": XLSX_TEMPLATE_TXT  # 添加对应的配置文件
    }
}
DEFAULT_MERGE_MODE = "docx"  # 默认合成模式

# 关于信息
ABOUT_TEXT = f"""
FileMerge v{VERSION}

功能说明：
1. 支持Word和Excel模板文件合成
2. 支持自定义配置文件（txt格式）
3. 支持文件名格式和文件夹格式自定义
4. 支持快捷键操作

使用说明：
1. 选择默认模版（Word或Excel）
2. 选择配置文件和模板文件
3. 填写数据
4. 点击"合成"按钮生成文件

特色功能：
1. Word文档按入院日期自动分类存储
2. Excel文件按申请日期年月自动分类存储
3. 支持jinja2模板语法，可实现条件判断
4. 文件名和文件夹格式可在设置菜单中配置

注意事项：
* 当前版本主题切换功能存在问题，暂时禁用
* 请使用默认主题以确保最佳体验

作者：{AUTHOR}
邮箱：{EMAIL}
"""

# 快捷键说明
SHORTCUTS_HELP = """
快捷键说明：

1. 文件操作
   {shortcut}+M : 合成文件
   {shortcut}+O : 打开最近生成的文件

2. 文件目录
   - 可通过"文件"菜单打开各类模板目录
   - 可通过"文件"菜单打开输出目录

3. 设置选项
   - Word设置：文件名格式、文件夹格式
   - Excel设置：文件名格式、文件夹格式

注：
1. {shortcut} 在 macOS 上是 Command，在 Windows/Linux 上是 Ctrl
2. 主题切换功能当前版本存在问题，暂时禁用
"""

# 日志文件
LOG_FILE = os.path.join(PROJECT_ROOT, "documents", "merge_log.txt") 

# 文件夹格式配置
FOLDER_FORMATS = {
    "docx": {
        "入院日期格式": "{{ 入院日期 }}"  # 格式如"2024年11月11日"
    },
    "excel": {
        "申请日期格式": "{{ 申请日期[:4] }}年{{ 申请日期[5:7] }}月"  # 格式如"2024年11月"
    }
}

# 文件名格式配置（使用jinja2语法）
FILENAME_FORMATS = {
    "docx": {
        "序号_姓名_手术眼别_模板名称": "{{ 序号 }}_{{ 姓名 }}_{{ 手术眼别 }}_{{ template_name }}",
        "序号_姓名_手术眼别_手术日期_模板名称": "{{ 序号 }}_{{ 姓名 }}_{{ 手术眼别 }}_{{ 手术日期 }}_{{ template_name }}"
    },
    "excel": {
        "申请日期_姓名_眼别_诊断_药品_针数": "{{ 申请日期 }}_{{ 姓名 }}_{{ 眼别 }}_{{ 备选诊断 }}_{{ 药品名称 }}_第{{ (地米植入累计|int + 阿柏康柏累计|int + 1)|string }}针",
        "姓名_眼别_诊断_药品_针数": "{{ 姓名 }}_{{ 眼别 }}_{{ 备选诊断 }}_{{ 药品名称 }}_第{{ (地米植入累计|int + 阿柏康柏累计|int + 1)|string }}针"
    }
}

# 默认输出目录
DEFAULT_OUTPUT_DIRS = {
    "docx": "手术日期",  # 使用手术日期作为子目录
    "excel": "申请日期"   # 使用申请日期作为子目录
} 