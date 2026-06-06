# -*- coding: utf-8 -*-
"""
Windows打包后的启动入口
自动打开浏览器并启动Flask服务
"""

import os
import sys
import webbrowser
from threading import Timer

# 获取打包后的资源路径
def get_resource_path(relative_path):
    """获取打包后的资源路径"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller打包后的临时目录
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def open_browser():
    """打开浏览器"""
    webbrowser.open('http://127.0.0.1:5001/')


if __name__ == '__main__':
    # 确保工作目录正确
    if hasattr(sys, '_MEIPASS'):
        os.chdir(sys._MEIPASS)

    # 延迟打开浏览器
    Timer(1.5, open_browser).start()

    # 导入并启动Flask应用
    from app import app
    app.run(host='0.0.0.0', port=5001, debug=False)
