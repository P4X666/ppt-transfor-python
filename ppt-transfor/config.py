"""
项目配置文件

集中管理所有可配置参数，便于维护和调整
"""

import os

# 项目根目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 默认输出目录
DEFAULT_OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# 默认资源目录
DEFAULT_ASSETS_DIR = os.path.join(BASE_DIR, "assets")

# 图片资源目录
DEFAULT_IMAGES_DIR = os.path.join(DEFAULT_ASSETS_DIR, "images")

# JSON输出文件名
DEFAULT_JSON_FILENAME = "output.json"

# 还原PPT文件名
DEFAULT_RECONSTRUCTED_FILENAME = "reconstructed.pptx"

# 日志配置
LOG_CONFIG = {
    "level": "INFO",
    "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    "date_format": "%Y-%m-%d %H:%M:%S"
}

# 图片提取配置
IMAGE_CONFIG = {
    "save_base64": True,      # 是否在JSON中嵌入base64编码的图片
    "save_file": True,        # 是否保存图片文件到磁盘
    "deduplicate": True,      # 是否对图片进行去重
    "supported_formats": ["png", "jpg", "jpeg", "gif", "bmp", "tiff"]
}

# 颜色配置
COLOR_CONFIG = {
    "format": "hex",          # 颜色输出格式: hex/rgb
    "include_theme_color": True,  # 是否包含主题颜色信息
    "include_brightness": True    # 是否包含亮度调整信息
}

# 文本提取配置
TEXT_CONFIG = {
    "extract_runs": True,     # 是否提取文本run级别的格式
    "extract_paragraph_format": True,  # 是否提取段落格式
    "preserve_empty_paragraphs": False  # 是否保留空段落
}

# PPT尺寸配置 (EMU单位: 1英寸 = 914400 EMU)
SLIDE_DIMENSIONS = {
    "standard_4_3": {
        "width": 9144000,     # 10英寸
        "height": 6858000     # 7.5英寸
    },
    "widescreen_16_9": {
        "width": 12192000,    # 13.333英寸
        "height": 6858000     # 7.5英寸
    }
}
