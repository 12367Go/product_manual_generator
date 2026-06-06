# -*- coding: utf-8 -*-
"""
PyInstaller 打包脚本
将Flask应用打包为Windows可执行文件

使用方法:
    python build.py

打包完成后，dist/产品手册生成器/ 目录即为最终交付文件夹
"""

import os
import sys
import subprocess
import shutil

# 项目根目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def check_pyinstaller():
    """检查是否安装了PyInstaller"""
    try:
        import PyInstaller
        return True
    except ImportError:
        return False


def install_pyinstaller():
    """安装PyInstaller"""
    print("正在安装 PyInstaller...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    print("PyInstaller 安装完成")


def clean_build():
    """清理之前的构建文件"""
    dirs_to_remove = ['build', 'dist']
    for d in dirs_to_remove:
        path = os.path.join(BASE_DIR, d)
        if os.path.exists(path):
            print(f"清理 {d}/ ...")
            shutil.rmtree(path)

    # 清理spec文件
    for f in os.listdir(BASE_DIR):
        if f.endswith('.spec'):
            os.remove(os.path.join(BASE_DIR, f))
            print(f"清理 {f}")


def create_spec():
    """创建PyInstaller的spec配置文件"""
    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

import os

block_cipher = None

# 项目根目录
base_dir = os.path.abspath(r'{BASE_DIR}')

# 收集所有需要打包的数据文件
datas = [
    # 模板文件
    (os.path.join(base_dir, 'templates'), 'templates'),
    # 静态文件(CSS/JS)
    (os.path.join(base_dir, 'static'), 'static'),
    # 配置文件
    (os.path.join(base_dir, 'config.py'), '.'),
]

# 添加core模块作为隐藏导入
hiddenimports = [
    'core.ppt_parser',
    'core.image_handler',
    'core.price_matcher',
    'core.export_manager',
    'config',
    'flask',
    'jinja2',
    'werkzeug',
    'pptx',
    'PIL',
    'pandas',
    'openpyxl',
    'xlrd',
    'numpy',
    'lxml',
]

a = Analysis(
    [os.path.join(base_dir, 'run.py')],
    pathex=[base_dir],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='产品手册生成器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不显示控制台窗口
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(base_dir, 'static', 'favicon.ico') if os.path.exists(os.path.join(base_dir, 'static', 'favicon.ico')) else None,
)
'''

    spec_path = os.path.join(BASE_DIR, 'product_manual.spec')
    with open(spec_path, 'w', encoding='utf-8') as f:
        f.write(spec_content)

    return spec_path


def build():
    """执行打包"""
    print("=" * 50)
    print("产品手册生成器 - PyInstaller 打包")
    print("=" * 50)

    # 检查/安装PyInstaller
    if not check_pyinstaller():
        install_pyinstaller()

    # 清理旧构建
    clean_build()

    # 创建spec文件
    spec_path = create_spec()
    print(f"创建配置文件: {spec_path}")

    # 执行打包
    print("\n开始打包...")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        spec_path,
        "--clean",
        "--noconfirm",
    ]

    try:
        subprocess.check_call(cmd)
        print("\n" + "=" * 50)
        print("打包成功！")
        print("=" * 50)
        print(f"\n输出目录: {os.path.join(BASE_DIR, 'dist', '产品手册生成器')}")
        print("\n请将整个 'dist/产品手册生成器' 文件夹压缩后发送给朋友")
        print("朋友解压后双击 '产品手册生成器.exe' 即可运行")
    except subprocess.CalledProcessError as e:
        print(f"\n打包失败: {e}")
        sys.exit(1)


if __name__ == '__main__':
    build()
