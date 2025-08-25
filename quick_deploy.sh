#!/bin/bash
# 快捷部署命令

echo "🚀 快速部署到 GitHub Pages..."

# 如果是首次运行
if [ ! -d ".git" ]; then
    echo "📂 初始化Git仓库..."
    git init
    git add .
    git commit -m "Initial commit"
    git branch -M main
    git remote add origin https://github.com/kongruijie123/excel-viewer.git
    git push -u origin main

    echo "✅ 首次上传完成！"
    echo "🌐 请访问 https://github.com/kongruijie123/excel-viewer/settings/pages 启用 GitHub Pages"
else
    # 更新部署
    echo "📝 重新生成静态网站..."
    python3 excel_viewer_static.py

    git add .
    git commit -m "更新数据 $(date '+%Y-%m-%d %H:%M:%S')"
    git push

    echo "✅ 更新完成！"
fi

echo "🌐 访问地址: https://kongruijie123.github.io/excel-viewer"
