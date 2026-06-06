# -*- coding: utf-8 -*-
"""
PyInstaller build script
Build Flask app into Windows executable

Usage:
    python build.py

After build, dist/product_manual_generator/ is the final deliverable folder
"""

import os
import sys
import subprocess
import shutil

# Project root directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def check_pyinstaller():
    """Check if PyInstaller is installed"""
    try:
        import PyInstaller
        return True
    except ImportError:
        return False


def install_pyinstaller():
    """Install PyInstaller"""
    print("Installing PyInstaller...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    print("PyInstaller installed")


def clean_build():
    """Clean previous build files"""
    dirs_to_remove = ['build', 'dist']
    for d in dirs_to_remove:
        path = os.path.join(BASE_DIR, d)
        if os.path.exists(path):
            print(f"Cleaning {d}/ ...")
            shutil.rmtree(path)

    # Clean spec files
    for f in os.listdir(BASE_DIR):
        if f.endswith('.spec'):
            os.remove(os.path.join(BASE_DIR, f))
            print(f"Cleaning {f}")


def create_spec():
    """Create PyInstaller spec file"""
    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

import os

block_cipher = None

# Project root directory
base_dir = os.path.abspath(r'{BASE_DIR}')

# Collect all data files needed
datas = [
    # Templates
    (os.path.join(base_dir, 'templates'), 'templates'),
    # Static files (CSS/JS)
    (os.path.join(base_dir, 'static'), 'static'),
    # Config file
    (os.path.join(base_dir, 'config.py'), '.'),
]

# Hidden imports
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
    name='product_manual_generator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
'''

    spec_path = os.path.join(BASE_DIR, 'product_manual.spec')
    with open(spec_path, 'w', encoding='utf-8') as f:
        f.write(spec_content)

    return spec_path


def build():
    """Run build"""
    print("=" * 50)
    print("Product Manual Generator - PyInstaller Build")
    print("=" * 50)

    # Check/install PyInstaller
    if not check_pyinstaller():
        install_pyinstaller()

    # Clean old builds
    clean_build()

    # Create spec file
    spec_path = create_spec()
    print(f"Created spec: {spec_path}")

    # Run build
    print("\nStarting build...")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        spec_path,
        "--clean",
        "--noconfirm",
    ]

    try:
        subprocess.check_call(cmd)
        print("\n" + "=" * 50)
        print("Build successful!")
        print("=" * 50)
        print(f"\nOutput: {os.path.join(BASE_DIR, 'dist', 'product_manual_generator')}")
        print("\nZip the 'dist/product_manual_generator' folder and send to your friend")
        print("Friend can double-click 'product_manual_generator.exe' to run")
    except subprocess.CalledProcessError as e:
        print(f"\nBuild failed: {e}")
        sys.exit(1)


if __name__ == '__main__':
    build()
