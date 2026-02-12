#!/bin/bash

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "================================================"
echo "   水质检测报告验证分析系统"
echo "================================================"
echo ""

# 检查 Python
python_version=$(python3 --version 2>&1 | awk '{print $2}')
echo "Python 版本: $python_version"

# 检查并创建虚拟环境
if [ ! -d "venv" ]; then
    echo "正在创建虚拟环境..."
    python3 -m venv venv
    echo "虚拟环境创建完成"
fi

# 激活虚拟环境
source venv/bin/activate

# 安装依赖
pip install -q -r requirements.txt

echo ""
echo "开始分析报告文件..."
echo "================================================"
echo ""

# 运行分析脚本（可通过参数指定扫描目录）
python3 analyze_reports.py "$@"
