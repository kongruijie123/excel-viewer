
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
