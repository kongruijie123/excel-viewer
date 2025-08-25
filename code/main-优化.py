# ===== 方案1: 生成静态HTML文件 + 免费托管 =====

import json
import os
from datetime import datetime


def generate_static_website():
    """将Excel数据转换为静态网站"""

    # 1. 读取Excel文件并转换为JSON
    def convert_excel_to_static():
        """转换Excel为静态JSON数据"""
        excel_data = {}
        static_data = {
            'files': {},
            'generated_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'version': '1.0'
        }

        # 扫描Excel文件
        excel_folder = "excel_files"
        if os.path.exists(excel_folder):
            for filename in os.listdir(excel_folder):
                if filename.endswith(('.xlsx', '.xls')):
                    file_path = os.path.join(excel_folder, filename)
                    try:
                        import pandas as pd
                        df = pd.read_excel(file_path, sheet_name=None)

                        static_data['files'][filename] = {}
                        for sheet_name, sheet_df in df.items():
                            sheet_df = sheet_df.fillna('')
                            static_data['files'][filename][sheet_name] = {
                                'data': sheet_df.to_dict('records'),
                                'columns': list(sheet_df.columns),
                                'row_count': len(sheet_df),
                                'col_count': len(sheet_df.columns)
                            }
                    except Exception as e:
                        print(f"处理文件 {filename} 时出错: {e}")

        return static_data

    # 2. 生成静态HTML文件
    def create_static_html(data):
        """创建静态HTML查看器"""

        # 主页HTML
        index_html = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel表格查看器 - 静态版</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            font-family: 'Arial', sans-serif;
        }}
        .container {{ padding-top: 2rem; }}
        .card {{
            border: none;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            background: white;
        }}
        .card-header {{
            background: linear-gradient(45deg, #28a745, #20c997);
            color: white;
            border-radius: 15px 15px 0 0 !important;
            text-align: center;
            padding: 1.5rem;
        }}
        .file-item {{
            padding: 1.5rem;
            margin: 1rem 0;
            background: #f8f9fa;
            border-radius: 10px;
            border-left: 4px solid #28a745;
        }}
        .sheet-info {{
            background: #fff;
            padding: 0.8rem;
            margin: 0.5rem 0;
            border-radius: 8px;
            border-left: 3px solid #007bff;
            cursor: pointer;
        }}
        .sheet-info:hover {{ background: #f8f9fa; }}
        .btn-view {{
            background: linear-gradient(45deg, #007bff, #0056b3);
            border: none;
            color: white;
            padding: 0.5rem 1.5rem;
            border-radius: 25px;
        }}
        .static-badge {{
            background: linear-gradient(45deg, #ffc107, #ff8c00);
            color: white;
            padding: 0.5rem 1rem;
            border-radius: 20px;
            font-size: 0.9em;
            display: inline-block;
            margin-bottom: 1rem;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="row justify-content-center">
            <div class="col-md-10">
                <div class="card">
                    <div class="card-header">
                        <h1><i class="fas fa-table"></i> Excel表格查看器</h1>
                        <div class="static-badge">
                            <i class="fas fa-bolt"></i> 静态版本 - 永久访问
                        </div>
                        <p class="mb-0">数据生成时间: {data['generated_time']}</p>
                    </div>
                    <div class="card-body">
                        <div id="file-list">
                            <!-- 文件列表将由JavaScript动态生成 -->
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- 模态框显示表格数据 -->
    <div class="modal fade" id="dataModal" tabindex="-1">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalTitle">表格数据</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div id="tableContainer" style="max-height: 70vh; overflow: auto;">
                        <!-- 表格数据 -->
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // 数据嵌入到HTML中
        const excelData = {json.dumps(data, ensure_ascii=False, indent=2)};

        // 生成文件列表
        function generateFileList() {{
            const container = document.getElementById('file-list');
            const files = excelData.files;

            if (Object.keys(files).length === 0) {{
                container.innerHTML = `
                    <div class="text-center py-5">
                        <i class="fas fa-folder-open fa-3x text-muted mb-3"></i>
                        <h4>暂无Excel文件</h4>
                        <p class="text-muted">请重新生成静态网站</p>
                    </div>
                `;
                return;
            }}

            let html = '';
            for (const [filename, sheets] of Object.entries(files)) {{
                html += `
                    <div class="file-item">
                        <div class="d-flex justify-content-between align-items-start">
                            <div class="flex-grow-1">
                                <h4><i class="fas fa-file-excel text-success"></i> ${{filename}}</h4>
                                <p class="text-muted mb-2">
                                    <i class="fas fa-layer-group"></i> 工作表数量: ${{Object.keys(sheets).length}}
                                </p>
                                <div class="sheets-preview">
                `;

                for (const [sheetName, sheetData] of Object.entries(sheets)) {{
                    html += `
                        <div class="sheet-info" onclick="showSheet('${{filename}}', '${{sheetName}}')">
                            <div class="d-flex justify-content-between align-items-center">
                                <div>
                                    <strong><i class="fas fa-table text-primary"></i> ${{sheetName}}</strong>
                                    <span class="text-muted ms-2">
                                        ${{sheetData.row_count}} 行 × ${{sheetData.col_count}} 列
                                    </span>
                                </div>
                                <small class="text-muted">点击查看</small>
                            </div>
                        </div>
                    `;
                }}

                html += `
                                </div>
                            </div>
                        </div>
                    </div>
                `;
            }}

            container.innerHTML = html;
        }}

        // 显示工作表数据
        function showSheet(filename, sheetName) {{
            const sheetData = excelData.files[filename][sheetName];
            const modalTitle = document.getElementById('modalTitle');
            const tableContainer = document.getElementById('tableContainer');

            modalTitle.textContent = `${{filename}} - ${{sheetName}}`;

            if (sheetData.data.length === 0) {{
                tableContainer.innerHTML = `
                    <div class="text-center py-5">
                        <i class="fas fa-exclamation-triangle fa-3x text-muted"></i>
                        <h4>工作表为空</h4>
                    </div>
                `;
            }} else {{
                let tableHtml = `
                    <table class="table table-striped table-hover">
                        <thead class="table-dark">
                            <tr>
                                <th>#</th>
                `;

                // 表头
                sheetData.columns.forEach(col => {{
                    tableHtml += `<th>${{col}}</th>`;
                }});

                tableHtml += '</tr></thead><tbody>';

                // 数据行
                sheetData.data.forEach((row, index) => {{
                    tableHtml += `<tr><td>${{index + 1}}</td>`;
                    sheetData.columns.forEach(col => {{
                        tableHtml += `<td>${{row[col] || ''}}</td>`;
                    }});
                    tableHtml += '</tr>';
                }});

                tableHtml += '</tbody></table>';
                tableContainer.innerHTML = tableHtml;
            }}

            // 显示模态框
            const modal = new bootstrap.Modal(document.getElementById('dataModal'));
            modal.show();
        }}

        // 页面加载后生成文件列表
        document.addEventListener('DOMContentLoaded', generateFileList);
    </script>
</body>
</html>
        """

        return index_html

    # 执行生成过程
    print("🔄 正在生成静态网站...")

    # 创建输出目录
    static_dir = "static_website"
    if not os.path.exists(static_dir):
        os.makedirs(static_dir)

    # 转换数据
    data = convert_excel_to_static()

    # 生成HTML
    html_content = create_static_html(data)

    # 写入文件
    with open(os.path.join(static_dir, 'index.html'), 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"✅ 静态网站已生成到: {static_dir}/index.html")
    print(f"📊 包含 {len(data['files'])} 个Excel文件")

    return static_dir


# ===== 方案2: 免费静态网站托管平台 =====

def deploy_to_static_hosting():
    """部署到免费静态托管平台"""

    platforms = {
        "GitHub Pages": {
            "url": "https://pages.github.com/",
            "steps": [
                "1. 创建 GitHub 仓库",
                "2. 上传 static_website 文件夹内容",
                "3. 启用 GitHub Pages",
                "4. 访问 https://用户名.github.io/仓库名"
            ],
            "pros": "完全免费，自定义域名支持",
            "限制": "公开仓库，100GB流量/月"
        },
        "Netlify": {
            "url": "https://www.netlify.com/",
            "steps": [
                "1. 注册 Netlify 账户",
                "2. 拖拽 static_website 文件夹到网站",
                "3. 获得随机域名如 amazing-name-123.netlify.app",
                "4. 可绑定自定义域名"
            ],
            "pros": "即拖即用，HTTPS自动配置",
            "限制": "100GB带宽/月"
        },
        "Vercel": {
            "url": "https://vercel.com/",
            "steps": [
                "1. 注册 Vercel 账户",
                "2. 连接 GitHub 或直接拖拽上传",
                "3. 获得 .vercel.app 域名",
                "4. 支持自动部署"
            ],
            "pros": "性能极佳，全球CDN",
            "限制": "100GB带宽/月"
        },
        "Firebase Hosting": {
            "url": "https://firebase.google.com/",
            "steps": [
                "1. 创建 Firebase 项目",
                "2. npm install -g firebase-tools",
                "3. firebase init hosting",
                "4. firebase deploy"
            ],
            "pros": "Google基础设施，速度快",
            "限制": "10GB存储，360MB/天流量"
        }
    }

    print("\n🌐 免费静态托管平台对比:")
    print("=" * 60)

    for name, info in platforms.items():
        print(f"\n📋 {name}")
        print(f"网址: {info['url']}")
        print(f"优点: {info['pros']}")
        print(f"限制: {info['限制']}")
        print("部署步骤:")
        for step in info['steps']:
            print(f"   {step}")

    return platforms


# ===== 方案3: 自动化部署脚本 =====

def create_auto_deploy_script():
    """创建自动化部署脚本"""

    # GitHub Pages 自动部署
    github_workflow = """
# .github/workflows/deploy.yml
name: Deploy Static Site

on:
  push:
    branches: [ main ]

  # 允许手动触发
  workflow_dispatch:

jobs:
  deploy:
    runs-on: ubuntu-latest

    steps:
    - uses: actions/checkout@v2

    - name: Setup Python
      uses: actions/setup-python@v2
      with:
        python-version: '3.9'

    - name: Install dependencies
      run: |
        pip install pandas openpyxl xlrd

    - name: Generate static site
      run: |
        python generate_static.py

    - name: Deploy to GitHub Pages
      uses: peaceiris/actions-gh-pages@v3
      with:
        github_token: ${{ secrets.GITHUB_TOKEN }}
        publish_dir: ./static_website
"""

    # 一键部署脚本
    deploy_script = """
#!/bin/bash
# 一键部署脚本 deploy.sh

echo "🚀 开始生成和部署静态网站..."

# 生成静态网站
python3 excel_viewer_static.py

# 询问用户选择部署平台
echo "请选择部署平台:"
echo "1. GitHub Pages"
echo "2. Netlify (拖拽上传)"
echo "3. Vercel (拖拽上传)"

read -p "请选择 (1-3): " choice

case $choice in
    1)
        echo "📂 GitHub Pages 部署步骤:"
        echo "1. 创建 GitHub 仓库"
        echo "2. git init && git add . && git commit -m 'Initial commit'"
        echo "3. git remote add origin YOUR_REPO_URL"
        echo "4. git push -u origin main"
        echo "5. 在仓库设置中启用 GitHub Pages"
        ;;
    2)
        echo "🌐 Netlify 部署:"
        echo "1. 访问 https://app.netlify.com/drop"
        echo "2. 拖拽 static_website 文件夹到页面"
        echo "3. 等待部署完成，获得访问链接"
        ;;
    3)
        echo "⚡ Vercel 部署:"
        echo "1. 访问 https://vercel.com/new"
        echo "2. 选择上传文件夹"
        echo "3. 拖拽 static_website 文件夹"
        echo "4. 点击 Deploy"
        ;;
esac

echo "✅ 静态网站文件已准备完成！"
echo "📁 位置: ./static_website/"
"""

    # Windows批处理版本
    windows_script = """
@echo off
echo 🚀 开始生成和部署静态网站...

python excel_viewer_static.py

echo.
echo 请选择部署平台:
echo 1. GitHub Pages
echo 2. Netlify (拖拽上传)
echo 3. Vercel (拖拽上传)
echo.

set /p choice="请选择 (1-3): "

if "%choice%"=="1" (
    echo 📂 GitHub Pages 部署步骤:
    echo 1. 创建 GitHub 仓库
    echo 2. 将 static_website 文件夹内容上传到仓库
    echo 3. 在仓库设置中启用 GitHub Pages
)

if "%choice%"=="2" (
    echo 🌐 Netlify 部署:
    echo 1. 访问 https://app.netlify.com/drop
    echo 2. 拖拽 static_website 文件夹到页面
    echo 3. 等待部署完成，获得访问链接
)

if "%choice%"=="3" (
    echo ⚡ Vercel 部署:
    echo 1. 访问 https://vercel.com/new
    echo 2. 选择上传文件夹
    echo 3. 拖拽 static_website 文件夹
    echo 4. 点击 Deploy
)

echo.
echo ✅ 静态网站文件已准备完成！
echo 📁 位置: ./static_website/

pause
"""

    return github_workflow, deploy_script, windows_script


# ===== 主程序 =====

def main_static():
    """静态部署主程序"""

    print("🎯 Excel静态网站生成器")
    print("=" * 50)
    print("将Excel文件转换为静态网站，无需运行程序即可访问")
    print()

    # 生成静态网站
    static_dir = generate_static_website()

    # 显示部署选项
    print("\n🌐 部署选项:")
    platforms = deploy_to_static_hosting()

    # 创建部署脚本
    print("\n📝 创建部署脚本...")
    github_workflow, deploy_script, windows_script = create_auto_deploy_script()

    # 保存脚本文件
    with open("deploy.sh", "w", encoding="utf-8") as f:
        f.write(deploy_script)

    with open("deploy.bat", "w", encoding="utf-8") as f:
        f.write(windows_script)

    with open(".github_workflow_deploy.yml", "w", encoding="utf-8") as f:
        f.write(github_workflow)

    print("✅ 部署脚本已生成:")
    print("   - deploy.sh (Linux/Mac)")
    print("   - deploy.bat (Windows)")
    print("   - .github_workflow_deploy.yml (GitHub Actions)")

    print(f"\n🎉 完成！静态网站已生成")
    print(f"📁 位置: {static_dir}/")
    print(f"🌐 本地预览: 打开 {static_dir}/index.html")
    print(f"☁️ 在线部署: 将文件夹上传到任何静态托管平台")

    # 生成二维码（使用本地文件地址作为示例）
    import qrcode
    local_file_path = os.path.abspath(os.path.join(static_dir, 'index.html'))
    file_url = f"file:///{local_file_path}"

    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
    qr.add_data("https://your-deployed-site.netlify.app")  # 示例URL
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")
    img.save("static_qr_code.png")

    print(f"📱 二维码已生成: static_qr_code.png")
    print(f"💡 部署后将二维码中的URL替换为实际部署地址")


if __name__ == "__main__":
    main_static()