# -*- coding: utf-8 -*-
"""
价格匹配模块
负责解析用户上传的价格表（CSV/Excel），并根据Excel中的款号顺序匹配对应的产品图片和价格

核心规则：
1. 以Excel表格中的排序为准
2. 根据Excel中的款号查找对应图片（支持多种匹配方式）
3. 找不到图片的款号，图片路径留空（后续PPT中该位置不插入图片，但保留款号和价格）
4. 找不到价格的款号，价格显示为N/A
"""

import os
import pandas as pd

# 允许上传的价格表文件扩展名
ALLOWED_PRICE_EXTENSIONS = {'csv', 'xlsx', 'xls'}


class PriceMatcher:
    """价格匹配器"""
    
    def __init__(self):
        """初始化价格匹配器"""
        self.price_data = {}  # 存储价格数据 {filename: price}
        self.raw_data = None  # 原始DataFrame
        self.ordered_codes = []  # Excel中的款号顺序列表
    
    def load_price_file(self, file_path):
        """
        加载价格表文件
        
        Args:
            file_path: 价格表文件路径（CSV/Excel）
            
        Returns:
            bool: 是否加载成功
        """
        try:
            file_ext = os.path.splitext(file_path)[1].lower()
            
            if file_ext == '.csv':
                # 读取CSV文件
                self.raw_data = pd.read_csv(file_path, header=None)
            elif file_ext in ('.xlsx', '.xls'):
                # 读取Excel文件
                self.raw_data = pd.read_excel(file_path, header=None)
            else:
                print(f"不支持的文件格式: {file_ext}")
                return False
            
            # 解析价格数据
            self._parse_price_data()
            return True
            
        except Exception as e:
            print(f"加载价格表失败: {e}")
            return False
    
    def _parse_price_data(self):
        """
        解析价格数据，构建款号-价格映射，并保留Excel原始顺序
        
        假设表格结构：
        - 第一行可能是表头（如：款号, 价格）
        - 从第二行开始是数据
        - 第一列是款号
        - 第二列是价格
        """
        if self.raw_data is None or self.raw_data.empty:
            return
        
        self.ordered_codes = []
        
        # 检查第一行是否是表头
        start_row = 0
        if len(self.raw_data) > 0:
            first_col_value = self._extract_value(self.raw_data.iloc[0, 0])
            if first_col_value and ('款号' in first_col_value or 'code' in first_col_value.lower()):
                start_row = 1  # 跳过表头行
        
        # 遍历数据行
        for idx in range(start_row, len(self.raw_data)):
            row = self.raw_data.iloc[idx]
            
            # 跳过空行
            if row.isna().all():
                continue
            
            # 获取第一列的值（款号）
            code_value = self._extract_value(row.iloc[0])
            if not code_value:
                continue
            
            # 获取第二列的值（价格）
            price_value = self._extract_price(row.iloc[1] if len(row) > 1 else None)
            
            # 存储到字典并记录顺序
            self.ordered_codes.append(code_value)
            if price_value is not None:
                self.price_data[code_value] = price_value
    
    def _extract_value(self, value):
        """
        提取单元格的值
        """
        if pd.isna(value):
            return None
        
        # 转换为字符串并清理
        str_value = str(value).strip()
        
        # 如果是浮点数但表示整数（如 81061.0），转换为整数
        try:
            float_val = float(str_value)
            if float_val.is_integer():
                return str(int(float_val))
        except ValueError:
            pass
        
        return str_value
    
    def _extract_price(self, value):
        """
        提取价格值
        """
        if pd.isna(value):
            return None
        
        str_value = str(value).strip()
        
        # 去除价格符号和空格
        str_value = str_value.replace('¥', '').replace('$', '').replace(',', '').strip()
        
        try:
            # 尝试转换为数值
            price = float(str_value)
            # 如果是整数价格，返回整数
            if price.is_integer():
                return int(price)
            return price
        except ValueError:
            return None
    
    def _find_image_for_code(self, code, image_paths):
        """
        根据款号查找对应的图片路径
        
        匹配规则（按优先级）：
        1. 完全匹配（款号 == 文件名不含扩展名）
        2. 去除前导零匹配
        3. 包含关系匹配（款号包含于文件名或文件名包含于款号）
        
        Args:
            code: 款号
            image_paths: 所有图片路径列表
            
        Returns:
            str: 匹配到的图片路径，找不到返回None
        """
        # 规则1：完全匹配
        for img_path in image_paths:
            filename = os.path.basename(img_path)
            name_without_ext = os.path.splitext(filename)[0]
            if code == name_without_ext:
                return img_path
        
        # 规则2：去除前导零匹配
        stripped_code = code.lstrip('0')
        for img_path in image_paths:
            filename = os.path.basename(img_path)
            name_without_ext = os.path.splitext(filename)[0]
            if stripped_code == name_without_ext.lstrip('0'):
                return img_path
        
        # 规则3：包含关系匹配
        for img_path in image_paths:
            filename = os.path.basename(img_path)
            name_without_ext = os.path.splitext(filename)[0]
            if code in name_without_ext or name_without_ext in code:
                return img_path
        
        return None
    
    def match_all_products(self, image_paths):
        """
        按Excel顺序生成产品列表
        
        核心逻辑：
        1. 遍历Excel中的款号（保持原始顺序）
        2. 为每个款号查找对应图片
        3. 获取对应价格
        4. 找不到图片的款号，image_path为None（后续PPT中不插入图片，但保留款号和价格）
        
        注意：不在Excel中的图片不会加入列表
        
        Args:
            image_paths: 所有上传的图片路径列表
            
        Returns:
            list: 产品列表，每个产品包含{'image_path', 'code', 'price'}
        """
        products = []
        
        # 按Excel顺序遍历款号
        for code in self.ordered_codes:
            # 查找对应图片
            image_path = self._find_image_for_code(code, image_paths)
            
            # 获取价格
            price = self.price_data.get(code, 'N/A')
            
            products.append({
                'image_path': image_path,  # 可能为None
                'code': code,
                'price': price
            })
        
        return products
    
    def get_price_summary(self):
        """
        获取价格匹配摘要
        """
        total = len(self.ordered_codes)
        matched = len([c for c in self.ordered_codes if c in self.price_data])
        return {
            'total_records': total,
            'matched_prices': matched,
            'sample_data': dict(list(self.price_data.items())[:5])  # 前5条示例
        }
