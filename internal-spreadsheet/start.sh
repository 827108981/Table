#!/bin/bash

# 内部在线表格系统启动脚本

echo "🚀 正在启动内部在线表格系统..."

# 进入项目目录
cd "$(dirname "$0")"

# 检查依赖是否安装
if ! python -c "import flask" 2>/dev/null; then
    echo "⚠️  检测到未安装依赖，正在安装..."
    pip install -r requirements.txt
fi

# 启动服务
echo "✅ 服务启动成功！"
echo ""
echo "📍 访问地址：http://localhost:5000"
echo "👤 默认账号：admin / admin123"
echo ""
echo "⚠️  首次登录后请立即修改密码！"
echo ""
echo "按 Ctrl+C 停止服务"
echo ""

python backend/app.py
