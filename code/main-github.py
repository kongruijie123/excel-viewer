# ===== Mac GitHub Pages 完整部署指南 =====

import os
import json
import pandas as pd
import qrcode
from datetime import datetime
import subprocess
# 确保输出到docs文件夹
output_dir = "docs"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

def create_static_excel_viewer():
    """创建静态Excel查看器"""

    print("🔄 正在生成静态网站...")

    # 1. 读取Excel文件并转换为JSON
    excel_data = {}
    excel_folder = "excel_files"

    if not os.path.exists(excel_folder):
        os.makedirs(excel_folder)
        print(f"❌ 请先将Excel文件放入 {excel_folder} 文件夹")
        return None

    # 在处理Excel文件的部分，将原来的try-except修改为：
    for filename in os.listdir(excel_folder):
        if filename.endswith(('.xlsx', '.xls')):
            file_path = os.path.join(excel_folder, filename)
            try:
                df = pd.read_excel(file_path, sheet_name=None)
                excel_data[filename] = {}

                print(f"开始处理文件: {filename}")
                print(f"工作表列表: {list(df.keys())}")

                for sheet_name, sheet_df in df.items():
                    print(f"  处理工作表: {sheet_name}")
                    sheet_df = sheet_df.fillna('')

                    # 检查数据转换是否成功
                    try:
                        data_records = sheet_df.to_dict('records')
                        excel_data[filename][sheet_name] = {
                            'data': data_records,
                            'columns': list(sheet_df.columns),
                            'row_count': len(sheet_df),
                            'col_count': len(sheet_df.columns)
                        }
                        print(f"    成功转换 {len(data_records)} 行数据")
                    except Exception as sheet_error:
                        print(f"    工作表 {sheet_name} 转换失败: {sheet_error}")

                print(f"文件 {filename} 最终包含 {len(excel_data[filename])} 个工作表")

            except Exception as e:
                print(f"处理文件 {filename} 时出错: {e}")
                print(f"错误类型: {type(e).__name__}")
    # 2. 创建输出目录
    output_dir = "docs"  # GitHub Pages 推荐使用 docs 文件夹
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    # 在生成 html_content 之前添加这些检查
    print("\n=== JSON序列化前检查 ===")
    print(f"excel_data包含的文件: {list(excel_data.keys())}")

    # 检查JSON序列化是否成功
    try:
        json_str = json.dumps(excel_data, ensure_ascii=False, indent=2)
        print("JSON序列化成功")

        # 检查序列化后的JSON字符串是否包含所有文件
        if "简化板.xlsx" in json_str:
            print("✅ JSON中包含简化板.xlsx")
        else:
            print("❌ JSON中缺少简化板.xlsx")

    except Exception as json_error:
        print(f"JSON序列化失败: {json_error}")

    # 然后继续原有的 html_content = f"""...""" 代码
    # 3. 生成完整的HTML页面
    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel数据查看器</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            font-family: 'Arial', sans-serif;
        }}
        .main-container {{ padding: 2rem 0; }}
        .card {{
            border: none;
            border-radius: 20px;
            box-shadow: 0 15px 35px rgba(0,0,0,0.1);
            background: rgba(255,255,255,0.95);
            backdrop-filter: blur(10px);
        }}
        .card-header {{
            background: linear-gradient(45deg, #28a745, #20c997);
            color: white;
            border-radius: 20px 20px 0 0 !important;
            text-align: center;
            padding: 2rem;
            border: none;
        }}
        .file-card {{
            background: #f8f9fa;
            border-radius: 15px;
            padding: 1.5rem;
            margin: 1rem 0;
            border-left: 5px solid #28a745;
            transition: all 0.3s ease;
        }}
        .file-card:hover {{
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.1);
            background: #e9ecef;
        }}
        .sheet-badge {{
            background: linear-gradient(45deg, #007bff, #0056b3);
            color: white;
            padding: 0.5rem 1rem;
            border-radius: 20px;
            cursor: pointer;
            transition: all 0.3s ease;
            display: inline-block;
            margin: 0.25rem;
        }}
        .sheet-badge:hover {{
            transform: scale(1.05);
            box-shadow: 0 5px 15px rgba(0,123,255,0.3);
        }}
        .table-container {{
            max-height: 70vh;
            overflow: auto;
            border-radius: 10px;
            background: white;
        }}
        .table {{
            margin-bottom: 0;
        }}
        .table th {{
            background: #f8f9fa;
            border-top: none;
            position: sticky;
            top: 0;
            z-index: 10;
            font-weight: 600;
        }}
        .table tbody tr:hover {{
            background-color: #e3f2fd;
        }}
        .stats-badge {{
            background: rgba(255,255,255,0.2);
            color: white;
            padding: 0.5rem 1rem;
            border-radius: 15px;
            font-size: 0.9em;
            display: inline-block;
            margin: 0.25rem;
        }}
        .github-badge {{
            position: fixed;
            top: 20px;
            right: 20px;
            background: #333;
            color: white;
            padding: 0.5rem 1rem;
            border-radius: 20px;
            text-decoration: none;
            font-size: 0.9em;
            z-index: 1000;
        }}
        .github-badge:hover {{
            background: #555;
            color: white;
        }}
    </style>
</head>
<body>
    <!-- GitHub 角标 -->
    <a href="#" class="github-badge" id="github-link">
        <i class="fab fa-github"></i> View on GitHub
    </a>

    <div class="container main-container">
        <div class="row justify-content-center">
            <div class="col-lg-10">
                <div class="card">
                    <div class="card-header">
                        <h1><i class="fas fa-table"></i> Excel数据查看器</h1>
                        <div>
                            <span class="stats-badge">
                                <i class="fas fa-database"></i> {len(excel_data)} 个文件
                            </span>
                            <span class="stats-badge">
                                <i class="fas fa-clock"></i> {datetime.now().strftime('%Y-%m-%d %H:%M')}
                            </span>
                        </div>
                    </div>
                    <div class="card-body">
                        <div id="file-list">
                            <!-- 文件列表 -->
                        </div>

                        <!-- 数据显示区域 -->
                        <div id="data-display" style="display: none;">
                            <div class="d-flex justify-content-between align-items-center mb-3">
                                <h4 id="current-title"></h4>
                                <button class="btn btn-outline-secondary" onclick="showFileList()">
                                    <i class="fas fa-arrow-left"></i> 返回列表
                                </button>
                            </div>
                            <div id="table-container" class="table-container">
                                <!-- 表格数据 -->
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // Excel数据嵌入
        const excelData = {json.dumps(excel_data, ensure_ascii=False, indent=2)};

        // 设置GitHub链接
        document.getElementById('github-link').href = window.location.origin + window.location.pathname;

        // 生成文件列表
        function generateFileList() {{
            const container = document.getElementById('file-list');
            let html = '';

            for (const [filename, sheets] of Object.entries(excelData)) {{
                const totalRows = Object.values(sheets).reduce((sum, sheet) => sum + sheet.row_count, 0);

                html += `
                    <div class="file-card">
                        <div class="d-flex justify-content-between align-items-start mb-3">
                            <div>
                                <h4 class="mb-1">
                                    <i class="fas fa-file-excel text-success"></i> 
                                    ${{filename}}
                                </h4>
                                <small class="text-muted">
                                    <i class="fas fa-layer-group"></i> ${{Object.keys(sheets).length}} 个工作表
                                    <i class="fas fa-chart-bar ms-3"></i> 共 ${{totalRows}} 行数据
                                </small>
                            </div>
                        </div>
                        <div class="sheet-badges">
                `;

                for (const [sheetName, sheetData] of Object.entries(sheets)) {{
                    html += `
                        <span class="sheet-badge" onclick="showSheet('${{filename}}', '${{sheetName}}')">
                            <i class="fas fa-table"></i> ${{sheetName}}
                            <small>(${{sheetData.row_count}}×${{sheetData.col_count}})</small>
                        </span>
                    `;
                }}

                html += `
                        </div>
                    </div>
                `;
            }}

            container.innerHTML = html || `
                <div class="text-center py-5">
                    <i class="fas fa-folder-open fa-4x text-muted mb-3"></i>
                    <h4>暂无数据</h4>
                    <p class="text-muted">请检查Excel文件是否正确处理</p>
                </div>
            `;
        }}

        // 显示工作表数据
        function showSheet(filename, sheetName) {{
            const sheetData = excelData[filename][sheetName];

            // 更新标题
            document.getElementById('current-title').innerHTML = `
                <i class="fas fa-file-excel text-success"></i> ${{filename}} 
                <i class="fas fa-angle-right mx-2"></i> 
                <i class="fas fa-table text-primary"></i> ${{sheetName}}
            `;

            // 生成表格
            let tableHtml = `
                <div class="table-responsive">
                    <table class="table table-striped table-hover">
                        <thead class="table-dark">
                            <tr>
                                <th style="min-width: 60px;">#</th>
            `;

            // 表头
            sheetData.columns.forEach(col => {{
                tableHtml += `<th style="min-width: 120px;">${{col || '未命名列'}}</th>`;
            }});

            tableHtml += '</tr></thead><tbody>';

            // 数据行
            if (sheetData.data.length === 0) {{
                tableHtml += `
                    <tr>
                        <td colspan="${{sheetData.columns.length + 1}}" class="text-center py-4">
                            <i class="fas fa-info-circle text-muted"></i> 此工作表暂无数据
                        </td>
                    </tr>
                `;
            }} else {{
                sheetData.data.forEach((row, index) => {{
                    tableHtml += `<tr><td class="fw-bold">${{index + 1}}</td>`;
                    sheetData.columns.forEach(col => {{
                        const cellValue = row[col] || '';
                        tableHtml += `<td title="${{cellValue}}">${{cellValue}}</td>`;
                    }});
                    tableHtml += '</tr>';
                }});
            }}

            tableHtml += '</tbody></table></div>';

            document.getElementById('table-container').innerHTML = tableHtml;

            // 切换显示
            document.getElementById('file-list').style.display = 'none';
            document.getElementById('data-display').style.display = 'block';

            // 滚动到顶部
            window.scrollTo(0, 0);
        }}

        // 返回文件列表
        function showFileList() {{
            document.getElementById('file-list').style.display = 'block';
            document.getElementById('data-display').style.display = 'none';
            window.scrollTo(0, 0);
        }}

        // 页面加载完成后初始化
        document.addEventListener('DOMContentLoaded', function() {{
            generateFileList();
            console.log('Excel数据查看器已加载完成');
            console.log('包含文件:', Object.keys(excelData));
        }});
    </script>
</body>
</html>"""

    # 4. 写入HTML文件
    index_path = os.path.join(output_dir, 'index.html')
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"✅ 静态网站已生成: {index_path}")
    print(f"📊 包含 {len(excel_data)} 个Excel文件")

    return output_dir


def create_github_deployment_guide():
    """创建GitHub部署指南和脚本"""

    # 创建README文件
    readme_content = """# Excel数据查看器

这是一个静态网站，用于查看Excel表格数据。

## 访问地址

🔗 **在线访问**: [https://你的用户名.github.io/仓库名](https://你的用户名.github.io/仓库名)

## 功能特点

- 📱 支持手机和电脑访问
- 🔍 可查看多个Excel文件的所有工作表
- 📊 数据以表格形式展示
- 🌐 完全静态，无需服务器运行
- ⚡ 响应速度快，支持全球访问

## 数据更新

要更新Excel数据，请：

1. 更新 `excel_files/` 文件夹中的Excel文件
2. 运行 `python generate_static.py`
3. 提交并推送更改到GitHub

---

*最后更新: """ + datetime.now().strftime('%Y-%m-%d %H:%M:%S') + "*"

    # 创建.gitignore文件
    gitignore_content = """# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
env/
venv/
.env

# macOS
.DS_Store
.AppleDouble
.LSOverride

# Excel临时文件
~$*.xlsx
~$*.xls

# IDE
.vscode/
.idea/

# 日志文件
*.log

# 原始Excel文件（可选，如果不想公开Excel源文件）
# excel_files/
"""

    # Mac部署脚本
    deploy_script = """#!/bin/bash

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
"""

    # 保存文件
    files_to_create = {
        "README.md": readme_content,
        ".gitignore": gitignore_content,
        "deploy.sh": deploy_script
    }

    for filename, content in files_to_create.items():
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(content)

    # 使deploy.sh可执行
    os.chmod("deploy.sh", 0o755)

    print("✅ GitHub部署文件已创建:")
    for filename in files_to_create.keys():
        print(f"   - {filename}")


def generate_qr_code_for_github(username, repo_name):
    """为GitHub Pages生成二维码"""

    github_url = f"https://{username}.github.io/{repo_name}"

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=12,
        border=4,
    )
    qr.add_data(github_url)
    qr.make(fit=True)

    # 生成普通二维码
    img = qr.make_image(fill_color="black", back_color="white")
    img.save("github_pages_qr.png")

    # 生成高清打印版
    qr_hd = qrcode.QRCode(
        version=2,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=20,
        border=6,
    )
    qr_hd.add_data(github_url)
    qr_hd.make(fit=True)

    img_hd = qr_hd.make_image(fill_color="black", back_color="white")
    img_hd.save("github_pages_qr_hd.png")

    print(f"📱 GitHub Pages 二维码已生成:")
    print(f"   - github_pages_qr.png")
    print(f"   - github_pages_qr_hd.png")
    print(f"🌐 二维码指向: {github_url}")

    return github_url


def main_github_pages():
    """GitHub Pages 完整部署主程序"""

    print("🎯 GitHub Pages 部署向导")
    print("=" * 50)

    # 第一步：生成静态网站
    print("\n📋 第一步：生成静态网站")
    docs_dir = create_static_excel_viewer()

    if not docs_dir:
        print("❌ 静态网站生成失败，请检查Excel文件")
        return

    # 第二步：创建部署文件
    print("\n📋 第二步：创建GitHub部署文件")
    create_github_deployment_guide()

    # 第三步：获取用户信息
    print("\n📋 第三步：GitHub仓库信息")
    print("请提供以下信息来完成部署：")

    username = input("🔸 你的GitHub用户名: ").strip()
    repo_name = input("🔸 仓库名称 (建议: excel-viewer): ").strip() or "excel-viewer"

    if not username:
        print("❌ 用户名不能为空")
        return

    # 第四步：生成二维码
    print("\n📋 第四步：生成访问二维码")
    github_url = generate_qr_code_for_github(username, repo_name)

    # 第五步：显示完整操作步骤
    print(f"\n🎉 准备完成！请按以下步骤操作：")
    print("=" * 60)

    print(f"\n1️⃣ 创建GitHub仓库:")
    print(f"   • 访问 https://github.com/new")
    print(f"   • 仓库名: {repo_name}")
    print(f"   • 设为Public（公开）")
    print(f"   • 点击 Create repository")

    print(f"\n2️⃣ 上传代码到仓库:")
    print(f"   打开终端，在当前文件夹执行：")
    print(f"   git init")
    print(f"   git add .")
    print(f"   git commit -m 'Initial commit'")
    print(f"   git branch -M main")
    print(f"   git remote add origin https://github.com/{username}/{repo_name}.git")
    print(f"   git push -u origin main")

    print(f"\n3️⃣ 启用GitHub Pages:")
    print(f"   • 访问 https://github.com/{username}/{repo_name}/settings/pages")
    print(f"   • Source 选择 'Deploy from a branch'")
    print(f"   • Branch 选择 'main'")
    print(f"   • Folder 选择 '/docs'")
    print(f"   • 点击 Save")

    print(f"\n4️⃣ 访问你的网站:")
    print(f"   🌐 网址: {github_url}")
    print(f"   📱 二维码: github_pages_qr.png")
    print(f"   ⏱️ 等待1-5分钟生效")

    print(f"\n💡 更新数据的方法:")
    print(f"   • 修改 excel_files/ 中的Excel文件")
    print(f"   • 运行: python3 excel_viewer_static.py")
    print(f"   • 运行: ./deploy.sh （或手动git提交）")

    print(f"\n📞 如果遇到问题:")
    print(f"   • 确保仓库是Public")
    print(f"   • 检查 docs/index.html 是否存在")
    print(f"   • GitHub Pages 需要几分钟生效")

    # 创建快捷命令脚本
    quick_commands = f"""#!/bin/bash
# 快捷部署命令

echo "🚀 快速部署到 GitHub Pages..."

# 如果是首次运行
if [ ! -d ".git" ]; then
    echo "📂 初始化Git仓库..."
    git init
    git add .
    git commit -m "Initial commit"
    git branch -M main
    git remote add origin https://github.com/{username}/{repo_name}.git
    git push -u origin main

    echo "✅ 首次上传完成！"
    echo "🌐 请访问 https://github.com/{username}/{repo_name}/settings/pages 启用 GitHub Pages"
else
    # 更新部署
    echo "📝 重新生成静态网站..."
    python3 excel_viewer_static.py

    git add .
    git commit -m "更新数据 $(date '+%Y-%m-%d %H:%M:%S')"
    git push

    echo "✅ 更新完成！"
fi

echo "🌐 访问地址: {github_url}"
"""

    with open("quick_deploy.sh", 'w', encoding='utf-8') as f:
        f.write(quick_commands)
    os.chmod("quick_deploy.sh", 0o755)

    print(f"\n🚀 创建了快捷部署脚本: ./quick_deploy.sh")
    print(f"   以后只需运行这个脚本即可快速更新！")


if __name__ == "__main__":
    main_github_pages()