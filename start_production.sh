#!/bin/bash

# PPT to Images API Server 生产环境启动脚本
# Ubuntu 服务器使用

clear
echo "🚀 PPT to Images API Server (Production)"
printf '%.0s=' {1..60} && echo
echo ""

# 设置生产环境变量
export API_BASE_URL="http://157.245.153.64:4000"
export PORT=4000
export HOST="0.0.0.0"

echo "📍 环境配置:"
echo "   API_BASE_URL: $API_BASE_URL"
echo "   PORT: $PORT"
echo "   HOST: $HOST"
echo ""

# 检查是否在虚拟环境中
if [ -z "$VIRTUAL_ENV" ]; then
    echo "⚠️  未激活虚拟环境"
    
    # 尝试激活虚拟环境
    if [ -d "venv" ]; then
        echo "   正在激活 venv..."
        source venv/bin/activate
    else
        echo "   建议创建虚拟环境:"
        echo "   python3 -m venv venv"
        echo "   source venv/bin/activate"
        echo ""
    fi
fi

# 检查 LibreOffice
if ! command -v libreoffice &> /dev/null && ! command -v soffice &> /dev/null; then
    echo "⚠️  未检测到 LibreOffice"
    echo "   安装: sudo apt-get install libreoffice"
    echo ""
fi

# 检查 poppler
if ! command -v pdftoppm &> /dev/null; then
    echo "⚠️  未检测到 poppler-utils"
    echo "   安装: sudo apt-get install poppler-utils"
    echo ""
fi

# 检查依赖
if ! python3 -c "import fastapi" 2>/dev/null; then
    echo "⚠️  缺少依赖。正在安装..."
    pip install -r requirements.txt
fi

echo "✅ 依赖检查完成"
echo ""

printf '%.0s=' {1..60} && echo
echo ""
echo "🌐 服务地址: $API_BASE_URL"
echo "📚 API 文档: $API_BASE_URL/docs"
echo "🔍 健康检查: $API_BASE_URL/health"
echo ""
printf '%.0s=' {1..60} && echo
echo ""
echo "💡 使用提示:"
echo "  • 浏览器访问 Web 界面"
echo "  • 图片 URL 将使用配置的域名"
echo "  • 按 Ctrl+C 停止服务"
echo ""
printf '%.0s=' {1..60} && echo
echo ""

# 启动服务器
python3 api_server.py
