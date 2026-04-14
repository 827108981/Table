#!/bin/bash

echo "======================================================"
echo "迈瑞技术支持数据中心 - 地图精准定位功能"
echo "======================================================"
echo ""

# 检查Python
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误: 未找到 Python3"
    echo "请先安装 Python 3.8+"
    exit 1
fi

echo "✅ Python3: $(python3 --version)"
echo ""

# 检查依赖
echo "📦 检查依赖包..."
python3 -c "import openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "⚠️  检测到缺少依赖包"
    echo ""
    read -p "是否现在安装依赖包? (y/n): " install_deps
    if [ "$install_deps" = "y" ] || [ "$install_deps" = "Y" ]; then
        echo ""
        echo "正在安装依赖..."
        pip3 install -r requirements.txt
        if [ $? -ne 0 ]; then
            echo "❌ 依赖包安装失败"
            exit 1
        fi
        echo "✅ 依赖包安装完成"
    else
        echo "⏭️  跳过依赖安装（某些功能可能无法使用）"
    fi
else
    echo "✅ 依赖包已安装"
fi

echo ""
echo "======================================================"
echo "功能说明"
echo "======================================================"
echo ""
echo "本次更新解决了乡镇级别医院无法精准定位的问题"
echo ""
echo "✅ 已实现的功能:"
echo "  1. 百度地图API地理编码（需配置API Key）"
echo "  2. 区县级坐标映射（覆盖200+区县）"
echo "  3. 城市级坐标映射（覆盖100+城市）"
echo "  4. 智能多级降级策略"
echo ""
echo "📖 详细文档:"
echo "  - 百度地图API配置指南.md  (API申请教程)"
echo "  - 更新说明_地图精准定位.md  (功能说明)"
echo ""

echo "======================================================"
echo "配置百度地图API（可选，推荐）"
echo "======================================================"
echo ""
echo "如果不配置API，系统会使用本地坐标映射（精度较低）"
echo ""
read -p "是否现在配置百度地图API Key? (y/n): " config_api
if [ "$config_api" = "y" ] || [ "$config_api" = "Y" ]; then
    echo ""
    echo "请按照以下步骤配置:"
    echo "1. 访问: https://lbsyun.baidu.com/apiconsole/key"
    echo "2. 登录并创建应用（选择'服务端'）"
    echo "3. 复制访问应用（AK）"
    echo "4. 编辑 getHospital.py 文件"
    echo "5. 找到第 154 行: BAIDU_AK = \"你的百度地图AK\""
    echo "6. 替换为你的AK"
    echo ""
    echo "详细教程请查看: 百度地图API配置指南.md"
fi

echo ""
echo "======================================================"
echo "测试功能"
echo "======================================================"
echo ""
read -p "是否运行测试? (y/n): " run_test
if [ "$run_test" = "y" ] || [ "$run_test" = "Y" ]; then
    echo ""
    python3 test_simple.py
fi

echo ""
echo "======================================================"
echo "启动系统"
echo "======================================================"
echo ""
echo "运行以下命令启动完整系统:"
echo "  python3 main.py"
echo ""
echo "运行以下命令仅处理数据:"
echo "  python3 getHospital.py"
echo ""
echo "======================================================"
echo "感谢使用！"
echo "======================================================"
