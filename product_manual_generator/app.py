# -*- coding: utf-8 -*-
"""
Flask应用主入口
提供Web服务和API接口
"""

import os
import sys

# 修复Anaconda MKL库冲突问题（Mac上常见）
os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'

import webbrowser
import shutil
from threading import Timer
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import Config, OUTPUT_FOLDER, UPLOAD_FOLDER, AVAILABLE_FONTS, DEFAULT_STYLES, DEFAULT_IMAGE_SIZE
from core.ppt_parser import PPTTemplateParser, PPTGenerator
from core.image_handler import ImageHandler
from core.price_matcher import PriceMatcher
from core.export_manager import ExportManager

# 创建Flask应用
app = Flask(__name__)
app.config.from_object(Config)

# 设置Flask上传临时目录（解决某些环境下/tmp不可用的问题）
UPLOAD_TMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tmp')
os.makedirs(UPLOAD_TMP_DIR, exist_ok=True)
os.environ['TMPDIR'] = UPLOAD_TMP_DIR
os.environ['TEMP'] = UPLOAD_TMP_DIR
os.environ['TMP'] = UPLOAD_TMP_DIR

# 确保目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route('/')
def index():
    """主页面"""
    return render_template('index.html', fonts=AVAILABLE_FONTS)


@app.route('/api/upload/template', methods=['POST'])
def upload_template():
    """
    上传PPT模板文件
    
    Returns:
        JSON: 上传结果，包含模板结构信息
    """
    try:
        if 'template' not in request.files:
            return jsonify({'success': False, 'error': '未找到上传的文件'})
        
        file = request.files['template']
        if file.filename == '':
            return jsonify({'success': False, 'error': '文件名为空'})
        
        # 检查文件类型
        if not file.filename.endswith('.pptx'):
            return jsonify({'success': False, 'error': '仅支持.pptx格式的PPT文件'})
        
        # 保存文件
        template_path = os.path.join(UPLOAD_FOLDER, 'template.pptx')
        file.save(template_path)
        
        # 解析模板
        parser = PPTTemplateParser(template_path)
        
        # 返回模板信息
        template_info = {
            'slide_count': len(parser.prs.slides),
            'slide_width': parser.slide_width.inches,
            'slide_height': parser.slide_height.inches,
            'header_slide': 1,
            'content_template': 2,
            'footer_slide': len(parser.prs.slides),
            'table_rows': parser.template_structure['table_shape']['rows'] if parser.template_structure['table_shape'] else 0,
            'table_cols': parser.template_structure['table_shape']['cols'] if parser.template_structure['table_shape'] else 0,
            'product_positions': len([c for c in parser.template_structure['table_info'] if c['is_product_code_value']])
        }
        
        return jsonify({
            'success': True,
            'template_info': template_info
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/upload/images', methods=['POST'])
def upload_images():
    """
    上传图片文件和/或ZIP压缩包
    
    Returns:
        JSON: 上传结果，包含图片列表
    """
    try:
        handler = ImageHandler()
        
        # 获取上传的图片文件
        image_files = request.files.getlist('images')
        zip_file = request.files.get('zip')
        
        # 过滤有效的图片文件（排除空文件名）
        valid_image_files = []
        if image_files:
            for f in image_files:
                if f and f.filename and f.filename.strip():
                    valid_image_files.append(f)
        
        # 处理文件
        image_paths = handler.process_uploaded_files(
            files=valid_image_files if valid_image_files else None,
            zip_file=zip_file if zip_file and zip_file.filename else None
        )
        
        # 提取图片信息
        images_info = []
        for path in image_paths:
            info = handler.get_image_info(path)
            if info:
                images_info.append({
                    'filename': info['filename'],
                    'name_without_ext': handler.extract_filename_without_ext(path),
                    'width': info['width'],
                    'height': info['height'],
                    'format': info['format']
                })
        
        return jsonify({
            'success': True,
            'count': len(images_info),
            'images': images_info
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/upload/price', methods=['POST'])
def upload_price():
    """
    上传价格表文件（CSV/Excel）
    
    Returns:
        JSON: 上传结果，包含价格数据预览
    """
    try:
        if 'price' not in request.files:
            return jsonify({'success': False, 'error': '未找到上传的文件'})
        
        file = request.files['price']
        if file.filename == '':
            return jsonify({'success': False, 'error': '文件名为空'})
        
        # 保存文件
        price_path = os.path.join(UPLOAD_FOLDER, 'price_table' + os.path.splitext(file.filename)[1])
        file.save(price_path)
        
        # 加载价格数据
        matcher = PriceMatcher()
        success = matcher.load_price_file(price_path)
        
        if not success:
            return jsonify({'success': False, 'error': '无法解析价格表文件'})
        
        # 返回价格数据摘要
        summary = matcher.get_price_summary()
        
        return jsonify({
            'success': True,
            'total_records': summary['total_records'],
            'sample_data': summary['sample_data']
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/generate', methods=['POST'])
def generate_manual():
    """
    生成产品手册
    
    接收所有配置参数，生成PPT并导出为指定格式
    
    Returns:
        JSON: 生成结果，包含下载链接
    """
    try:
        data = request.get_json()
        
        # 获取配置参数
        series_name = data.get('series_name', '')
        series_font = data.get('series_font', DEFAULT_STYLES['series_name']['font'])
        series_size = int(data.get('series_size', DEFAULT_STYLES['series_name']['size']))
        series_bold = data.get('series_bold', DEFAULT_STYLES['series_name']['bold'])
        series_italic = data.get('series_italic', DEFAULT_STYLES['series_name']['italic'])
        
        product_info = data.get('product_info', '')
        info_font = data.get('info_font', DEFAULT_STYLES['product_info']['font'])
        info_size = int(data.get('info_size', DEFAULT_STYLES['product_info']['size']))
        info_bold = data.get('info_bold', DEFAULT_STYLES['product_info']['bold'])
        info_italic = data.get('info_italic', DEFAULT_STYLES['product_info']['italic'])
        
        image_width = float(data.get('image_width', DEFAULT_IMAGE_SIZE['width']))
        image_height = float(data.get('image_height', DEFAULT_IMAGE_SIZE['height']))
        
        export_formats = data.get('export_formats', ['pptx'])
        
        # 检查必要文件是否存在
        template_path = os.path.join(UPLOAD_FOLDER, 'template.pptx')
        if not os.path.exists(template_path):
            return jsonify({'success': False, 'error': '请先上传PPT模板'})
        
        # 加载图片
        handler = ImageHandler()
        image_folder = os.path.join(UPLOAD_FOLDER, 'images')
        image_paths = []
        if os.path.exists(image_folder):
            for f in sorted(os.listdir(image_folder)):
                if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp')):
                    image_paths.append(os.path.join(image_folder, f))
        
        if not image_paths:
            return jsonify({'success': False, 'error': '请先上传产品图片'})
        
        # 加载价格数据
        matcher = PriceMatcher()
        price_path = None
        for ext in ['.xlsx', '.xls', '.csv']:
            temp_path = os.path.join(UPLOAD_FOLDER, 'price_table' + ext)
            if os.path.exists(temp_path):
                price_path = temp_path
                break
        
        if price_path:
            matcher.load_price_file(price_path)
        
        # 按Excel顺序匹配产品（款号→图片→价格）
        products = matcher.match_all_products(image_paths)
        
        # 解析模板
        parser = PPTTemplateParser(template_path)
        
        # 生成PPT
        generator = PPTGenerator(parser)
        image_size = {'width': image_width, 'height': image_height}
        
        # 系列名称配置
        series_config = {
            'font': series_font,
            'size': series_size,
            'bold': series_bold,
            'italic': series_italic
        }
        
        # 产品信息配置
        info_config = {
            'font': info_font,
            'size': info_size,
            'bold': info_bold,
            'italic': info_italic
        }
        
        ppt_path = generator.generate(
            products=products,
            series_name=series_name,
            series_config=series_config,
            product_info=product_info,
            info_config=info_config,
            image_size=image_size
        )
        
        # 导出其他格式
        export_manager = ExportManager(OUTPUT_FOLDER)
        export_results = export_manager.export(ppt_path, export_formats)
        
        # 构建下载链接
        downloads = {}
        for fmt, path in export_results.items():
            if path and os.path.exists(path):
                filename = os.path.basename(path)
                downloads[fmt] = f'/api/download/{filename}'
        
        # 生成成功后清理上传的临时文件（保留生成的结果文件）
        _cleanup_uploads()

        return jsonify({
            'success': True,
            'message': '产品手册生成成功！',
            'downloads': downloads,
            'product_count': len(products)
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


def _cleanup_uploads():
    """
    清理上传的临时文件（图片、价格表、模板）
    保留 output 目录中的生成结果
    """
    import shutil
    cleaned = []

    # 清理图片目录
    image_folder = os.path.join(UPLOAD_FOLDER, 'images')
    if os.path.exists(image_folder):
        shutil.rmtree(image_folder)
        cleaned.append('images')

    # 清理价格表文件
    for ext in ['.xlsx', '.xls', '.csv']:
        price_path = os.path.join(UPLOAD_FOLDER, 'price_table' + ext)
        if os.path.exists(price_path):
            os.remove(price_path)
            cleaned.append('price_table' + ext)

    # 清理模板文件
    template_path = os.path.join(UPLOAD_FOLDER, 'template.pptx')
    if os.path.exists(template_path):
        os.remove(template_path)
        cleaned.append('template.pptx')

    if cleaned:
        print(f"已清理上传文件: {', '.join(cleaned)}")


@app.route('/api/download/<filename>')
def download_file(filename):
    """
    下载生成的文件
    
    Args:
        filename: 文件名
        
    Returns:
        File: 文件下载
    """
    try:
        return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/clear', methods=['POST'])
def clear_uploads():
    """
    清除所有上传的文件
    
    Returns:
        JSON: 操作结果
    """
    try:
        # 清除上传目录
        if os.path.exists(UPLOAD_FOLDER):
            shutil.rmtree(UPLOAD_FOLDER)
            os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        
        # 清除输出目录
        if os.path.exists(OUTPUT_FOLDER):
            shutil.rmtree(OUTPUT_FOLDER)
            os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        
        return jsonify({'success': True, 'message': '已清除所有文件'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


def open_browser():
    """自动打开浏览器"""
    webbrowser.open('http://127.0.0.1:5001/')


if __name__ == '__main__':
    # 延迟打开浏览器
    Timer(1.5, open_browser).start()
    
    # 启动Flask应用
    app.run(host='0.0.0.0', port=5001, debug=False)
