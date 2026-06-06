#!/bin/bash
# 产品手册生成器 - Mac/Linux 启动脚本

echo "=========================================="
echo "  产品手册生成器 启动中..."
echo "=========================================="

# 获取脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "错误：未找到 Python3，请先安装 Python3"
    exit 1
fi

echo "Python版本: $(python3 --version)"

# 检查虚拟环境
if [ ! -d "venv" ]; then
    echo "创建虚拟环境..."
    python3 -m venv venv
fi

# 激活虚拟环境
echo "激活虚拟环境..."
source venv/bin/activate

# 安装依赖
echo "检查并安装依赖..."
pip install -r requirements.txt --quiet

# 创建必要目录
mkdir -p static/uploads
mkdir -p output
mkdir -p tmp

# 启动应用
echo ""
echo "启动服务器..."
echo "请在浏览器中访问: http://127.0.0.1:5000"
echo ""

python3 app.py
