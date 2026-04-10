#!/bin/bash
echo "Mac版Excel比对工具打包脚本"
echo "========================="

# 检查Python3
if ! command -v python3 &> /dev/null
then
    echo "❌ 错误：未检测到Python3，请先安装Python"
    exit 1
fi

# 检查pyinstaller
if ! command -v pyinstaller &> /dev/null
then
    echo "⚠️  未检测到pyinstaller，正在安装..."
    pip3 install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
fi

echo "🚀 开始打包..."
pyinstaller -F -w --clean main.py

if [ -f "dist/main" ]
then
    echo ""
    echo "✅ 打包成功！"
    echo "📂 可执行文件路径：dist/main"
    echo "🔧 运行方式：双击dist/main 或者终端执行 ./dist/main"
    echo "⚠️  首次运行如果提示无法打开，请右键点击文件选择打开，或者在系统设置-隐私与安全性中允许运行"
else
    echo "❌ 打包失败，请查看错误信息"
fi