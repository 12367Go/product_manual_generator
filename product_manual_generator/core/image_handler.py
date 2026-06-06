# -*- coding: utf-8 -*-
"""
图片处理模块
负责处理用户上传的图片文件和ZIP压缩包
"""

import os
import zipfile
import shutil
from PIL import Image

# 允许上传的图片文件扩展名
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'}

# 上传文件夹路径
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'static', 'uploads')


class ImageHandler:
    """图片处理器"""
    
    def __init__(self, upload_folder=None):
        """
        初始化图片处理器
        
        Args:
            upload_folder: 上传文件存储目录
        """
        self.upload_folder = upload_folder or UPLOAD_FOLDER
        self.image_folder = os.path.join(self.upload_folder, 'images')
        os.makedirs(self.image_folder, exist_ok=True)
    
    def process_uploaded_files(self, files, zip_file=None):
        """
        处理用户上传的文件（图片文件和/或ZIP压缩包）
        
        Args:
            files: 上传的图片文件列表（Flask FileStorage对象列表）
            zip_file: 上传的ZIP文件（Flask FileStorage对象）
            
        Returns:
            list: 处理后的图片路径列表
        """
        all_image_paths = []
        
        # 清空之前的图片
        self._clear_image_folder()
        
        # 处理单个图片文件
        if files:
            for file in files:
                if file and self._allowed_image_file(file.filename):
                    image_path = self._save_image(file)
                    if image_path:
                        all_image_paths.append(image_path)
        
        # 处理ZIP压缩包
        if zip_file:
            zip_paths = self._extract_zip(zip_file)
            all_image_paths.extend(zip_paths)
        
        # 按文件名排序
        all_image_paths.sort(key=lambda x: os.path.basename(x))
        
        return all_image_paths
    
    def _clear_image_folder(self):
        """清空图片目录"""
        if os.path.exists(self.image_folder):
            shutil.rmtree(self.image_folder)
        os.makedirs(self.image_folder, exist_ok=True)
    
    def _allowed_image_file(self, filename):
        """
        检查文件是否是允许的图片格式
        支持大小写扩展名
        """
        if '.' not in filename:
            return False
        ext = filename.rsplit('.', 1)[1].lower()
        return ext in ALLOWED_IMAGE_EXTENSIONS
    
    def _save_image(self, file_storage):
        """
        保存上传的图片文件
        """
        try:
            filename = file_storage.filename
            safe_filename = os.path.basename(filename)
            
            file_path = os.path.join(self.image_folder, safe_filename)
            file_storage.save(file_path)
            
            return file_path
        except Exception as e:
            print(f"保存图片失败 {file_storage.filename}: {e}")
            return None
    
    def _extract_zip(self, zip_file_storage):
        """
        解压ZIP文件并提取其中的图片
        """
        extracted_paths = []
        
        try:
            temp_zip_path = os.path.join(self.upload_folder, 'temp.zip')
            zip_file_storage.save(temp_zip_path)
            
            with zipfile.ZipFile(temp_zip_path, 'r') as zip_ref:
                for file_info in zip_ref.infolist():
                    filename = file_info.filename
                    
                    if file_info.is_dir() or filename.startswith('__') or filename.startswith('.'):
                        continue
                    
                    if self._allowed_image_file(filename):
                        zip_ref.extract(file_info, self.image_folder)
                        
                        extracted_path = os.path.join(self.image_folder, filename)
                        
                        if os.path.dirname(extracted_path) != self.image_folder:
                            new_path = os.path.join(self.image_folder, os.path.basename(filename))
                            shutil.move(extracted_path, new_path)
                            extracted_path = new_path
                        
                        extracted_paths.append(extracted_path)
            
            if os.path.exists(temp_zip_path):
                os.remove(temp_zip_path)
            
            self._remove_empty_dirs(self.image_folder)
            
        except Exception as e:
            print(f"解压ZIP文件失败: {e}")
        
        return extracted_paths
    
    def _remove_empty_dirs(self, folder):
        """递归删除空目录"""
        for root, dirs, files in os.walk(folder, topdown=False):
            for dir_name in dirs:
                dir_path = os.path.join(root, dir_name)
                if not os.listdir(dir_path):
                    os.rmdir(dir_path)
    
    def get_image_info(self, image_path):
        """获取图片信息"""
        try:
            with Image.open(image_path) as img:
                return {
                    'width': img.width,
                    'height': img.height,
                    'format': img.format,
                    'mode': img.mode,
                    'filename': os.path.basename(image_path)
                }
        except Exception as e:
            print(f"获取图片信息失败 {image_path}: {e}")
            return None
    
    def extract_filename_without_ext(self, image_path):
        """提取图片文件名（不含扩展名），用于匹配价格表"""
        filename = os.path.basename(image_path)
        name_without_ext = os.path.splitext(filename)[0]
        return name_without_ext
