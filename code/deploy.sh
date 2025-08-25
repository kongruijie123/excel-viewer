#!/bin/bash

# GitHub Pages 自动部署脚本 (Mac版本)

echo "🚀 开始部署到GitHub Pages..."

# 检查是否在正确目录
if [ ! -f "excel_viewer.py" ]; then
    echo "❌ 请在项目根目录运行此脚本"
    exit 1
fi

# 检查是否安装了git
if ! command -v git &> /dev/null; then
    echo "❌ 请先安装Git: brew install git"
    exit 1
fi

# 生成静态网站
echo "📝 生成静态网站..."
python3 excel_viewer_static.py

if [ ! -d "docs" ]; then
    echo "❌ 静态网站生成失败"
    exit 1
fi

# 检查是否是git仓库
if [ ! -d ".git" ]; then
    echo "📂 初始化Git仓库..."
    git init
    git branch -M main
fi

# 添加文件
echo "📋 添加文件到Git..."
git add .
git commit -m "更新Excel数据 - $(date '+%Y-%m-%d %H:%M:%S')"

# 检查是否设置了远程仓库
if ! git remote get-url origin &> /dev/null; then
    echo "⚠️ 请设置GitHub远程仓库："
    echo "   git remote add origin https://github.com/你的用户名/你的仓库名.git"
    echo "   然后重新运行此脚本"
    exit 1
fi

# 推送到GitHub
echo "📤 推送到GitHub..."
git push -u origin main

echo "✅ 部署完成！"
echo "🌐 请访问 GitHub 仓库设置页面启用 GitHub Pages"
echo "📱 网站地址将是: https://你的用户名.github.io/仓库名"
