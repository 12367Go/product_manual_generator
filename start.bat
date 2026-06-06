@echo off
chcp 65001 >nul
REM 产品手册生成器 - Windows 启动脚本

echo ==========================================
echo   产品手册生成器 启动中...
echo ==========================================

REM 获取脚本所在目录
set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到 Python，请先安装 Python
    pause
    exit /b 1
)

echo Python版本:
python --version

REM 检查虚拟环境
if not exist "venv" (
    echo 创建虚拟环境...
    python -m venv venv
)

REM 激活虚拟环境
echo 激活虚拟环境...
call venv\Scripts\activate.bat

REM 安装依赖
echo 检查并安装依赖...
pip install -r requirements.txt --quiet

REM 创建必要目录
if not exist "static\uploads" mkdir static\uploads
if not exist "output" mkdir output
if not exist "tmp" mkdir tmp

REM 启动应用
echo.
echo 启动服务器...
echo 请在浏览器中访问: http://127.0.0.1:5000
echo.

python app.py

pause
