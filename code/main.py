import os
import pandas as pd
import qrcode
from flask import Flask, render_template, jsonify, redirect, url_for
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading
import time
from datetime import datetime
import socket
import json
import subprocess
import requests
import sys

app = Flask(__name__)

# 全局变量存储表格数据
excel_data = {}
last_modified = {}
public_url = None


class ExcelFileHandler(FileSystemEventHandler):
    """监控Excel文件变化的处理器"""

    def __init__(self, excel_folder):
        self.excel_folder = excel_folder

    def on_modified(self, event):
        if not event.is_directory and event.src_path.endswith(('.xlsx', '.xls')):
            print(f"检测到文件变化: {event.src_path}")
            load_excel_files()


def get_local_ip():
    """获取本机IP地址"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "localhost"


def check_ngrok_installed():
    """检查ngrok是否已安装"""
    try:
        subprocess.run(['ngrok', 'version'], capture_output=True, check=True)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False


def install_ngrok():
    """安装ngrok的说明"""
    print("\n" + "=" * 60)
    print("🔧 需要安装 ngrok 来实现公网访问")
    print("=" * 60)
    print("请按以下步骤安装ngrok：")
    print()
    print("1. 访问 https://ngrok.com/ 注册账号")
    print("2. 下载ngrok客户端")
    print("3. Windows: 下载ngrok.exe并放入PATH环境变量")
    print("   Mac: brew install ngrok")
    print("   Linux: sudo apt install ngrok (或从官网下载)")
    print("4. 运行: ngrok config add-authtoken YOUR_TOKEN")
    print("   (YOUR_TOKEN可在ngrok后台获取)")
    print()
    print("安装完成后重新运行此程序")
    print("=" * 60)
    return False


def start_ngrok(port):
    """启动ngrok隧道"""
    global public_url

    print("🚀 启动ngrok公网隧道...")

    try:
        # 启动ngrok进程
        ngrok_process = subprocess.Popen(
            ['ngrok', 'http', str(port), '--log=stdout'],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        # 等待ngrok启动
        time.sleep(5)

        # 获取ngrok公网地址
        try:
            response = requests.get('http://localhost:4040/api/tunnels')
            tunnels = response.json()

            if tunnels['tunnels']:
                public_url = tunnels['tunnels'][0]['public_url']
                print(f"✅ ngrok隧道已启动!")
                print(f"🌐 公网地址: {public_url}")
                return ngrok_process, public_url
            else:
                print("❌ 未能获取ngrok公网地址")
                return None, None

        except Exception as e:
            print(f"❌ 获取ngrok地址失败: {str(e)}")
            return None, None

    except Exception as e:
        print(f"❌ 启动ngrok失败: {str(e)}")
        return None, None


def load_excel_files(folder_path="excel_files"):
    """加载Excel文件并转换为数据"""
    global excel_data, last_modified

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"创建了文件夹: {folder_path}")
        return

    excel_data.clear()
    last_modified.clear()

    for filename in os.listdir(folder_path):
        if filename.endswith(('.xlsx', '.xls')):
            file_path = os.path.join(folder_path, filename)
            try:
                # 读取Excel文件
                df = pd.read_excel(file_path, sheet_name=None)  # 读取所有工作表

                excel_data[filename] = {}
                for sheet_name, sheet_df in df.items():
                    # 填充空值
                    sheet_df = sheet_df.fillna('')

                    # 转换为更结构化的数据
                    excel_data[filename][sheet_name] = {
                        'data': sheet_df.to_dict('records'),
                        'columns': list(sheet_df.columns),
                        'row_count': len(sheet_df),
                        'col_count': len(sheet_df.columns)
                    }

                # 记录文件修改时间
                last_modified[filename] = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime(
                    '%Y-%m-%d %H:%M:%S')
                print(f"成功加载: {filename}")

            except Exception as e:
                print(f"加载文件 {filename} 时出错: {str(e)}")


def generate_qr_code(url, filename="qr_code.png"):
    """生成二维码"""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")
    img.save(filename)
    print(f"📱 二维码已生成: {filename}")
    print(f"🔗 访问地址: {url}")

    # 同时生成一个高对比度的二维码用于打印
    qr_print = qrcode.QRCode(
        version=2,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=15,
        border=6,
    )
    qr_print.add_data(url)
    qr_print.make(fit=True)

    img_print = qr_print.make_image(fill_color="black", back_color="white")
    img_print.save("qr_code_print.png")
    print(f"🖨️ 高清打印版二维码: qr_code_print.png")


@app.route('/')
def index():
    """主页面"""
    return render_template('index.html',
                           excel_files=excel_data,
                           last_modified=last_modified,
                           public_url=public_url,
                           last_updated=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))


@app.route('/api/data')
def get_data():
    """API接口返回表格数据"""
    return jsonify({
        'data': excel_data,
        'last_modified': last_modified,
        'public_url': public_url,
        'timestamp': datetime.now().isoformat()
    })


@app.route('/view/<filename>')
def view_file(filename):
    """查看特定Excel文件的所有工作表"""
    if filename in excel_data:
        return render_template('table_view.html',
                               filename=filename,
                               sheets=excel_data[filename],
                               last_modified=last_modified.get(filename, '未知'),
                               public_url=public_url,
                               last_updated=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    else:
        return render_template('error.html', message=f"文件 '{filename}' 未找到"), 404


@app.route('/sheet/<filename>/<sheet_name>')
def view_sheet(filename, sheet_name):
    """查看特定工作表"""
    if filename in excel_data and sheet_name in excel_data[filename]:
        sheet_data = excel_data[filename][sheet_name]
        return render_template('sheet_view.html',
                               filename=filename,
                               sheet_name=sheet_name,
                               sheet_data=sheet_data,
                               public_url=public_url,
                               last_modified=last_modified.get(filename, '未知'))
    else:
        return render_template('error.html', message=f"工作表未找到"), 404


@app.route('/status')
def status():
    """服务状态页面"""
    return jsonify({
        'status': 'running',
        'files_count': len(excel_data),
        'public_url': public_url,
        'local_ip': get_local_ip(),
        'last_updated': datetime.now().isoformat()
    })


def create_templates():
    """创建HTML模板文件"""
    template_dir = "templates"
    if not os.path.exists(template_dir):
        os.makedirs(template_dir)

    # 主页面模板 (更新后包含公网访问信息)
    index_html = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel表格查看器 - 全球访问</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            font-family: 'Arial', sans-serif;
        }
        .container {
            padding-top: 2rem;
        }
        .card {
            border: none;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            transition: transform 0.3s ease;
        }
        .card:hover {
            transform: translateY(-5px);
        }
        .card-header {
            background: linear-gradient(45deg, #28a745, #20c997);
            color: white;
            border-radius: 15px 15px 0 0 !important;
            text-align: center;
            padding: 1.5rem;
        }
        .public-url-card {
            background: linear-gradient(45deg, #fd7e14, #e63946);
            color: white;
            border-radius: 15px;
            padding: 1.5rem;
            margin-bottom: 2rem;
            text-align: center;
        }
        .url-display {
            background: rgba(255,255,255,0.2);
            padding: 1rem;
            border-radius: 10px;
            font-family: monospace;
            font-size: 1.1em;
            word-break: break-all;
            margin: 1rem 0;
        }
        .file-list {
            max-height: 500px;
            overflow-y: auto;
        }
        .file-item {
            padding: 1.5rem;
            margin: 1rem 0;
            background: #f8f9fa;
            border-radius: 10px;
            transition: all 0.3s ease;
            border-left: 4px solid #28a745;
        }
        .file-item:hover {
            background: #e9ecef;
            transform: translateX(5px);
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }
        .sheet-info {
            background: #fff;
            padding: 0.8rem;
            margin: 0.5rem 0;
            border-radius: 8px;
            border-left: 3px solid #007bff;
        }
        .sheet-info:hover {
            background: #f8f9fa;
            cursor: pointer;
        }
        .update-time {
            color: #6c757d;
            font-size: 0.9em;
        }
        .refresh-btn {
            position: fixed;
            bottom: 30px;
            right: 30px;
            width: 60px;
            height: 60px;
            border-radius: 50%;
            background: linear-gradient(45deg, #fd7e14, #e63946);
            border: none;
            color: white;
            font-size: 1.5em;
            box-shadow: 0 5px 15px rgba(0,0,0,0.3);
            transition: all 0.3s ease;
        }
        .refresh-btn:hover {
            transform: scale(1.1);
        }
        .status-indicator {
            display: inline-block;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            background: #28a745;
            margin-right: 8px;
            animation: pulse 2s infinite;
        }
        .global-indicator {
            background: #dc3545;
            animation: pulse 2s infinite;
        }
        @keyframes pulse {
            0% { opacity: 1; }
            50% { opacity: 0.5; }
            100% { opacity: 1; }
        }
        .btn-view {
            background: linear-gradient(45deg, #007bff, #0056b3);
            border: none;
            color: white;
            padding: 0.5rem 1.5rem;
            border-radius: 25px;
            transition: all 0.3s ease;
        }
        .btn-view:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(0,123,255,0.3);
            color: white;
        }
        .stats {
            font-size: 0.85em;
            color: #6c757d;
        }
        .copy-btn {
            background: rgba(255,255,255,0.3);
            border: 1px solid rgba(255,255,255,0.5);
            color: white;
            border-radius: 20px;
            padding: 0.5rem 1rem;
            margin-left: 1rem;
            transition: all 0.3s ease;
        }
        .copy-btn:hover {
            background: rgba(255,255,255,0.5);
            color: white;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="row justify-content-center">
            <div class="col-md-10">
                {% if public_url %}
                <div class="public-url-card">
                    <h3><i class="fas fa-globe"></i> 全球公网访问地址</h3>
                    <p><span class="status-indicator global-indicator"></span>任何人扫描二维码都能访问</p>
                    <div class="url-display">
                        {{ public_url }}
                        <button class="copy-btn" onclick="copyUrl('{{ public_url }}')">
                            <i class="fas fa-copy"></i> 复制
                        </button>
                    </div>
                    <div class="mt-3">
                        <i class="fas fa-qrcode fa-2x"></i>
                        <p class="mb-0 mt-2">扫描 qr_code.png 或 qr_code_print.png</p>
                    </div>
                </div>
                {% endif %}

                <div class="card">
                    <div class="card-header">
                        <h1><i class="fas fa-table"></i> Excel表格查看器</h1>
                        <p class="mb-0">
                            {% if public_url %}
                            <span class="status-indicator global-indicator"></span>全球公网访问已启用
                            {% else %}
                            <span class="status-indicator"></span>本地局域网访问
                            {% endif %}
                        </p>
                    </div>
                    <div class="card-body">
                        {% if excel_files %}
                        <div class="file-list">
                            {% for filename, sheets in excel_files.items() %}
                            <div class="file-item">
                                <div class="d-flex justify-content-between align-items-start">
                                    <div class="flex-grow-1">
                                        <h4><i class="fas fa-file-excel text-success"></i> {{ filename }}</h4>
                                        <p class="stats mb-2">
                                            <i class="fas fa-clock"></i> 修改时间: {{ last_modified.get(filename, '未知') }}
                                            | <i class="fas fa-layer-group"></i> 工作表数量: {{ sheets|length }}
                                        </p>

                                        <div class="sheets-preview">
                                            {% for sheet_name, sheet_data in sheets.items() %}
                                            <div class="sheet-info" onclick="viewSheet('{{ filename }}', '{{ sheet_name }}')">
                                                <div class="d-flex justify-content-between align-items-center">
                                                    <div>
                                                        <strong><i class="fas fa-table text-primary"></i> {{ sheet_name }}</strong>
                                                        <span class="stats ms-2">
                                                            {{ sheet_data.row_count }} 行 × {{ sheet_data.col_count }} 列
                                                        </span>
                                                    </div>
                                                    <small class="text-muted">点击查看</small>
                                                </div>
                                            </div>
                                            {% endfor %}
                                        </div>
                                    </div>
                                    <div class="ms-3">
                                        <button class="btn btn-view" onclick="viewFile('{{ filename }}')">
                                            <i class="fas fa-eye"></i> 查看全部
                                        </button>
                                    </div>
                                </div>
                            </div>
                            {% endfor %}
                        </div>
                        {% else %}
                        <div class="text-center py-5">
                            <i class="fas fa-folder-open fa-3x text-muted mb-3"></i>
                            <h4>暂无Excel文件</h4>
                            <p class="text-muted">请将Excel文件放入 "excel_files" 文件夹中</p>
                            <button class="btn btn-primary" onclick="refreshData()">
                                <i class="fas fa-sync-alt"></i> 检查文件
                            </button>
                        </div>
                        {% endif %}

                        <div class="mt-4 text-center update-time">
                            <i class="fas fa-sync-alt"></i> 最后检查: {{ last_updated }}
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <button class="refresh-btn" onclick="refreshData()" title="刷新数据">
        <i class="fas fa-sync-alt" id="refresh-icon"></i>
    </button>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        function viewFile(filename) {
            console.log('查看文件:', filename);
            window.location.href = '/view/' + encodeURIComponent(filename);
        }

        function viewSheet(filename, sheetName) {
            console.log('查看工作表:', filename, sheetName);
            window.location.href = '/sheet/' + encodeURIComponent(filename) + '/' + encodeURIComponent(sheetName);
        }

        function copyUrl(url) {
            navigator.clipboard.writeText(url).then(function() {
                // 临时显示复制成功
                const btn = event.target.closest('.copy-btn');
                const originalText = btn.innerHTML;
                btn.innerHTML = '<i class="fas fa-check"></i> 已复制';
                setTimeout(() => {
                    btn.innerHTML = originalText;
                }, 2000);
            }).catch(function(err) {
                console.error('复制失败:', err);
                // 降级方案
                prompt('请手动复制这个地址:', url);
            });
        }

        function refreshData() {
            const icon = document.getElementById('refresh-icon');
            icon.style.animation = 'spin 1s linear infinite';

            fetch('/api/data')
                .then(response => response.json())
                .then(data => {
                    console.log('数据已更新', data.timestamp);
                    setTimeout(() => {
                        location.reload();
                    }, 1000);
                })
                .catch(error => {
                    console.error('刷新失败:', error);
                    icon.style.animation = '';
                });
        }

        // 自动刷新功能
        setInterval(refreshData, 30000); // 30秒自动刷新一次
    </script>

    <style>
        @keyframes spin {
            from { transform: rotate(0deg); }
            to { transform: rotate(360deg); }
        }
    </style>
</body>
</html>
    """

    # 表格查看页面模板（包含公网访问信息）
    table_view_html = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ filename }} - 完整查看</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 1rem 0;
        }
        .container-fluid {
            max-width: 95%;
        }
        .main-card {
            border: none;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            margin-bottom: 2rem;
            background: white;
        }
        .card-header {
            background: linear-gradient(45deg, #28a745, #20c997);
            color: white;
            border-radius: 15px 15px 0 0 !important;
            padding: 1.5rem;
        }
        .sheet-nav {
            background: #f8f9fa;
            padding: 1rem;
            border-bottom: 1px solid #dee2e6;
        }
        .sheet-tab {
            margin: 0 0.5rem;
            border-radius: 20px;
            border: 2px solid transparent;
            background: white;
            color: #495057;
            transition: all 0.3s ease;
        }
        .sheet-tab:hover {
            background: #e9ecef;
            transform: translateY(-2px);
        }
        .sheet-tab.active {
            background: linear-gradient(45deg, #28a745, #20c997);
            color: white;
            border-color: #20c997;
        }
        .table-container {
            max-height: 70vh;
            overflow: auto;
            border-radius: 0 0 15px 15px;
        }
        .table {
            font-size: 0.9em;
            margin-bottom: 0;
        }
        .table th {
            background: #f8f9fa;
            border-top: none;
            position: sticky;
            top: 0;
            z-index: 10;
            font-weight: 600;
        }
        .table tbody tr:hover {
            background: #f5f5f5;
        }
        .back-btn {
            position: fixed;
            top: 30px;
            left: 30px;
            z-index: 1000;
            border-radius: 50px;
            padding: 0.5rem 1.5rem;
        }
        .refresh-btn {
            position: fixed;
            bottom: 30px;
            right: 30px;
            width: 60px;
            height: 60px;
            border-radius: 50%;
            background: linear-gradient(45deg, #fd7e14, #e63946);
            border: none;
            color: white;
            font-size: 1.5em;
        }
        .sheet-info {
            background: #e3f2fd;
            padding: 1rem;
            border-radius: 10px;
            margin-bottom: 1rem;
        }
        .access-info {
            background: rgba(255,255,255,0.1);
            color: white;
            padding: 0.5rem 1rem;
            border-radius: 20px;
            font-size: 0.9em;
            display: inline-block;
        }
    </style>
</head>
<body>
    <button class="btn btn-primary back-btn" onclick="goBack()">
        <i class="fas fa-arrow-left"></i> 返回
    </button>

    <div class="container-fluid">
        <div class="main-card">
            <div class="card-header">
                <div class="row align-items-center">
                    <div class="col">
                        <h2><i class="fas fa-file-excel"></i> {{ filename }}</h2>
                        <p class="mb-2">文件修改时间: {{ last_modified }}</p>
                        {% if public_url %}
                        <div class="access-info">
                            <i class="fas fa-globe"></i> 全球公网访问
                        </div>
                        {% endif %}
                    </div>
                    <div class="col-auto">
                        <span class="badge bg-light text-dark fs-6">{{ sheets|length }} 个工作表</span>
                    </div>
                </div>
            </div>

            {% if sheets|length > 1 %}
            <div class="sheet-nav">
                <div class="d-flex flex-wrap justify-content-center">
                    {% for sheet_name in sheets.keys() %}
                    <button class="btn sheet-tab {% if loop.first %}active{% endif %}" 
                            onclick="showSheet('{{ sheet_name }}')">
                        <i class="fas fa-table"></i> {{ sheet_name }}
                        <span class="badge bg-secondary ms-2">{{ sheets[sheet_name].row_count }}</span>
                    </button>
                    {% endfor %}
                </div>
            </div>
            {% endif %}

            {% for sheet_name, sheet_data in sheets.items() %}
            <div class="sheet-content" id="sheet-{{ loop.index0 }}" 
                 style="{% if not loop.first %}display: none;{% endif %}">

                <div class="sheet-info">
                    <div class="row">
                        <div class="col-md-6">
                            <h5><i class="fas fa-table text-primary"></i> {{ sheet_name }}</h5>
                        </div>
                        <div class="col-md-6 text-end">
                            <span class="badge bg-primary">{{ sheet_data.row_count }} 行</span>
                            <span class="badge bg-success">{{ sheet_data.col_count }} 列</span>
                        </div>
                    </div>
                </div>

                <div class="table-container">
                    {% if sheet_data.data %}
                    <table class="table table-striped table-hover">
                        <thead>
                            <tr>
                                <th scope="col">#</th>
                                {% for col in sheet_data.columns %}
                                <th scope="col">{{ col }}</th>
                                {% endfor %}
                            </tr>
                        </thead>
                        <tbody>
                            {% for row in sheet_data.data %}
                            <tr>
                                <td>{{ loop.index }}</td>
                                {% for col in sheet_data.columns %}
                                <td>{{ row.get(col, '') }}</td>
                                {% endfor %}
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                    {% else %}
                    <div class="text-center py-5">
                        <i class="fas fa-exclamation-triangle fa-3x text-muted"></i>
                        <h4>工作表为空</h4>
                    </div>
                    {% endif %}
                </div>
            </div>
            {% endfor %}
        </div>
    </div>

    <button class="refresh-btn" onclick="location.reload()" title="刷新数据">
        <i class="fas fa-sync-alt"></i>
    </button>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        let currentSheet = 0;

        function showSheet(sheetName) {
            // 隐藏所有工作表
            document.querySelectorAll('.sheet-content').forEach((sheet, index) => {
                sheet.style.display = 'none';
            });

            // 显示选中的工作表
            const sheets = document.querySelectorAll('.sheet-content');
            const sheetIndex = Array.from(document.querySelectorAll('.sheet-tab')).findIndex(tab => 
                tab.textContent.trim().includes(sheetName)
            );

            if (sheetIndex >= 0) {
                sheets[sheetIndex].style.display = 'block';
                currentSheet = sheetIndex;
            }

            // 更新标签页状态
            document.querySelectorAll('.sheet-tab').forEach(tab => {
                tab.classList.remove('active');
            });
            event.target.classList.add('active');
        }

        function goBack() {
            window.location.href = '/';
        }

        // 自动刷新
        setInterval(() => {
            console.log('自动检查更新...');
            fetch('/api/data')
                .then(response => response.json())
                .then(data => {
                    // 检查文件是否有更新
                    const currentFile = '{{ filename }}';
                    if (data.last_modified[currentFile] !== '{{ last_modified }}') {
                        console.log('检测到文件更新，正在刷新页面...');
                        location.reload();
                    }
                })
                .catch(error => console.error('检查更新失败:', error));
        }, 10000); // 10秒检查一次更新
    </script>
</body>
</html>
    """

    # 单个工作表查看页面（包含公网访问信息）
    sheet_view_html = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ sheet_name }} - {{ filename }}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 1rem 0;
        }
        .container-fluid {
            max-width: 98%;
        }
        .main-card {
            border: none;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            margin-bottom: 2rem;
            background: white;
        }
        .card-header {
            background: linear-gradient(45deg, #007bff, #0056b3);
            color: white;
            border-radius: 15px 15px 0 0 !important;
            padding: 1.5rem;
        }
        .table-container {
            max-height: 80vh;
            overflow: auto;
            border-radius: 0 0 15px 15px;
        }
        .table {
            font-size: 0.85em;
            margin-bottom: 0;
        }
        .table th {
            background: #f8f9fa;
            border-top: none;
            position: sticky;
            top: 0;
            z-index: 10;
            font-weight: 600;
            white-space: nowrap;
        }
        .table tbody tr:nth-child(even) {
            background: #f9f9f9;
        }
        .table tbody tr:hover {
            background: #e3f2fd;
        }
        .table td {
            white-space: nowrap;
            max-width: 200px;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .back-btn {
            position: fixed;
            top: 30px;
            left: 30px;
            z-index: 1000;
            border-radius: 50px;
            padding: 0.5rem 1.5rem;
        }
        .stats {
            background: rgba(255,255,255,0.1);
            padding: 0.5rem 1rem;
            border-radius: 20px;
            display: inline-block;
        }
        .access-info {
            background: rgba(255,255,255,0.1);
            color: white;
            padding: 0.5rem 1rem;
            border-radius: 20px;
            font-size: 0.9em;
            display: inline-block;
            margin-left: 1rem;
        }
    </style>
</head>
<body>
    <button class="btn btn-primary back-btn" onclick="goBack()">
        <i class="fas fa-arrow-left"></i> 返回
    </button>

    <div class="container-fluid">
        <div class="main-card">
            <div class="card-header">
                <div class="row align-items-center">
                    <div class="col">
                        <h3><i class="fas fa-table"></i> {{ sheet_name }}</h3>
                        <p class="mb-2">来自文件: {{ filename }}</p>
                        <div class="stats">
                            <i class="fas fa-chart-bar"></i> 
                            {{ sheet_data.row_count }} 行 × {{ sheet_data.col_count }} 列
                        </div>
                        {% if public_url %}
                        <div class="access-info">
                            <i class="fas fa-globe"></i> 全球公网访问
                        </div>
                        {% endif %}
                    </div>
                    <div class="col-auto">
                        <small class="text-light">修改: {{ last_modified }}</small>
                    </div>
                </div>
            </div>

            <div class="table-container">
                {% if sheet_data.data %}
                <table class="table table-striped table-hover">
                    <thead>
                        <tr>
                            <th scope="col" style="min-width: 60px;">#</th>
                            {% for col in sheet_data.columns %}
                            <th scope="col" title="{{ col }}">{{ col }}</th>
                            {% endfor %}
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in sheet_data.data %}
                        <tr>
                            <td><strong>{{ loop.index }}</strong></td>
                            {% for col in sheet_data.columns %}
                            <td title="{{ row.get(col, '') }}">{{ row.get(col, '') }}</td>
                            {% endfor %}
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
                {% else %}
                <div class="text-center py-5">
                    <i class="fas fa-exclamation-triangle fa-3x text-muted"></i>
                    <h4>工作表为空</h4>
                    <p class="text-muted">此工作表没有数据</p>
                </div>
                {% endif %}
            </div>
        </div>
    </div>

    <script>
        function goBack() {
            window.location.href = '/';
        }

        // 双击单元格查看完整内容
        document.querySelectorAll('.table td').forEach(cell => {
            cell.addEventListener('dblclick', function() {
                const fullText = this.getAttribute('title') || this.textContent;
                if (fullText.trim()) {
                    alert(fullText);
                }
            });
        });
    </script>
</body>
</html>
    """

    # 错误页面模板
    error_html = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>错误 - Excel查看器</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .error-card {
            max-width: 500px;
            border: none;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        }
        .card-header {
            background: linear-gradient(45deg, #dc3545, #c82333);
            color: white;
            border-radius: 15px 15px 0 0 !important;
            text-align: center;
            padding: 2rem;
        }
    </style>
</head>
<body>
    <div class="card error-card">
        <div class="card-header">
            <i class="fas fa-exclamation-triangle fa-3x mb-3"></i>
            <h3>出错了</h3>
        </div>
        <div class="card-body text-center">
            <p class="lead">{{ message }}</p>
            <a href="/" class="btn btn-primary">
                <i class="fas fa-home"></i> 返回首页
            </a>
        </div>
    </div>
</body>
</html>
    """

    # 写入模板文件
    with open(os.path.join(template_dir, 'index.html'), 'w', encoding='utf-8') as f:
        f.write(index_html)

    with open(os.path.join(template_dir, 'table_view.html'), 'w', encoding='utf-8') as f:
        f.write(table_view_html)

    with open(os.path.join(template_dir, 'sheet_view.html'), 'w', encoding='utf-8') as f:
        f.write(sheet_view_html)

    with open(os.path.join(template_dir, 'error.html'), 'w', encoding='utf-8') as f:
        f.write(error_html)


def start_file_monitor(folder_path="excel_files"):
    """启动文件监控"""
    event_handler = ExcelFileHandler(folder_path)
    observer = Observer()
    observer.schedule(event_handler, folder_path, recursive=False)
    observer.start()
    print(f"开始监控文件夹: {folder_path}")
    return observer


def main():
    """主函数"""
    global public_url

    print("=== Excel二维码查看器 v3.0 - 全球公网版 ===")
    print("正在启动服务...")

    # 检查ngrok是否已安装
    if not check_ngrok_installed():
        if not install_ngrok():
            print("\n❌ 没有安装ngrok，将以局域网模式运行")
            print("如需公网访问，请安装ngrok后重新运行")
            time.sleep(3)

    # 创建必要的文件夹和模板
    create_templates()

    # 加载Excel文件
    load_excel_files()

    # 启动文件监控
    observer = start_file_monitor()

    # 获取本机IP和端口
    local_host = get_local_ip()
    port = 8000
    local_url = f"http://{local_host}:{port}"

    # 启动ngrok（如果可用）
    ngrok_process = None
    if check_ngrok_installed():
        try:
            ngrok_process, public_url = start_ngrok(port)
            if public_url:
                # 生成公网二维码
                generate_qr_code(public_url)
            else:
                # 降级到局域网模式
                print("⚠️ ngrok启动失败，降级到局域网模式")
                generate_qr_code(local_url)
        except Exception as e:
            print(f"⚠️ ngrok启动异常: {e}")
            print("降级到局域网模式")
            generate_qr_code(local_url)
    else:
        # 仅局域网模式
        generate_qr_code(local_url)

    # 显示访问信息
    print("\n" + "=" * 60)
    print("🚀 Excel二维码查看器已启动!")
    print("=" * 60)

    if public_url:
        print(f"🌐 全球公网地址: {public_url}")
        print(f"📱 扫描二维码即可全球访问")
        print(f"🔗 任何人都可以通过二维码查看你的Excel表格")
    else:
        print(f"🏠 本地局域网地址: {local_url}")
        print(f"📱 仅限同一WiFi网络内的设备访问")

    print(f"📁 Excel文件夹: excel_files/")
    print(f"📄 二维码文件: qr_code.png, qr_code_print.png")
    print(f"⚡ 文件监控: 已启用实时更新")
    print("\n按 Ctrl+C 停止服务")
    print("=" * 60)

    try:
        # 启动Flask应用
        app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
    except KeyboardInterrupt:
        print("\n正在停止服务...")
        observer.stop()
        if ngrok_process:
            try:
                ngrok_process.terminate()
                print("ngrok隧道已关闭")
            except:
                pass

    observer.join()
    print("服务已停止")


if __name__ == "__main__":
    main()