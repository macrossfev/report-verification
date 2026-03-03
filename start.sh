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

# 如果没有参数，显示帮助信息
if [ $# -eq 0 ]; then
    echo ""
    echo "用法: ./start.sh <模式> [其他选项]"
    echo ""
    echo "运行模式（必选其一）："
    echo "  -oridata       仅检查原始记录"
    echo "  -report        仅检查报告文件"
    echo "  -datareport    基于原始记录检查报告（交叉验证）"
    echo "  -public        基于报告核对公示表"
    echo ""
    echo "其他选项："
    echo "  -r <目录>      指定扫描目录"
    echo "  -o <文件>      自定义输出文件"
    echo "  -h             查看完整帮助"
    echo ""
    echo "示例："
    echo "  ./start.sh -oridata"
    echo "  ./start.sh -report -r /path/to/reports"
    echo "  ./start.sh -datareport"
    echo "  ./start.sh -public -r /path/to/Publicsheet"
    echo ""
    exit 0
fi

echo ""
echo "开始分析报告文件..."
echo "================================================"
echo ""

# 运行分析脚本，将所有参数传递给 Python
python3 analyze_reports.py "$@"
