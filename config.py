# -*- coding: utf-8 -*-
"""
全局配置文件
包含所有默认参数和常量定义
"""

import os

# 基础路径配置
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'static', 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'output')

# 确保目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Flask配置
class Config:
    """Flask应用配置类"""
    SECRET_KEY = 'product-manual-generator-secret-key'
    UPLOAD_FOLDER = UPLOAD_FOLDER
    OUTPUT_FOLDER = OUTPUT_FOLDER
    MAX_CONTENT_LENGTH = 500 * 1024 * 1024  # 最大上传500MB

# 默认样式配置
DEFAULT_STYLES = {
    'series_name': {
        'font': 'ABC Whyte',
        'size': 18,
        'bold': True,
        'italic': True
    },
    'product_info': {
        'font': 'ABC Whyte',
        'size': 9,
        'bold': False,
        'italic': False
    }
}

# 图片默认尺寸（厘米）
DEFAULT_IMAGE_SIZE = {
    'width': 5,
    'height': 5
}

# 字体列表（供用户选择）
AVAILABLE_FONTS = [
    '微软雅黑',
    '宋体',
    '黑体',
    '楷体',
    'Arial',
    'Times New Roman',
    'Copperplate Gothic Bold',
    'ABC Whyte',
    'Cadli',
    '思源黑体'
]

# PPT模板中的标签文本（用于识别替换位置）
TEMPLATE_LABELS = {
    'product_code': '款号',
    'retail_price': '零售价',
    'price_symbol': '¥',
    'price_symbol_fullwidth': '￥'
}
