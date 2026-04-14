@echo off
chcp 65001 >nul
echo ========================================
echo 迈瑞技术支持驾驶舱 - 打包工具
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到Python，请先安装Python 3.8+
    pause
    exit /b 1
)

echo [1/5] 安装必要的依赖包...
pip install playwright openpyxl "pyinstaller>=6.10.0,<6.19.0" -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo [错误] 依赖包安装失败，尝试使用官方源...
    pip install playwright openpyxl "pyinstaller>=6.10.0,<6.19.0"
    if errorlevel 1 (
        echo [错误] 依赖包安装仍然失败
        pause
        exit /b 1
    )
)
echo ✓ 依赖包安装完成
echo.

echo [2/5] 安装Playwright浏览器...
python -m playwright install msedge
if errorlevel 1 (
    echo [警告] Edge浏览器安装失败，尝试安装Chromium...
    python -m playwright install chromium
)
echo ✓ Playwright浏览器安装完成
echo.

echo [3/5] 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "*.spec" del /q *.spec
echo ✓ 清理完成
echo.

echo [4/5] 开始打包程序...
pyinstaller --onefile ^
    --name "迈瑞中国区技术支持数据中心" ^
    --add-data "dashboard.html;." ^
    --add-data "login.html;." ^
    --add-data "hospital_data.json;." ^
    --add-data "工单报表数据;工单报表数据" ^
    --hidden-import playwright ^
    --hidden-import openpyxl ^
    --hidden-import http.server ^
    --hidden-import socketserver ^
    --console ^
    main.py

if errorlevel 1 (
    echo [错误] 打包失败，请检查错误信息
    pause
    exit /b 1
)
echo ✓ 打包完成
echo.

echo [5/5] 创建发布文件夹...
if not exist "发布版本" mkdir "发布版本"
if exist "dist\迈瑞中国区技术支持数据中心.exe" (
    copy "dist\迈瑞中国区技术支持数据中心.exe" "发布版本\" >nul
    echo ✓ 可执行文件已复制到发布文件夹
) else (
    echo [错误] 未找到打包后的exe文件
    dir dist
    pause
    exit /b 1
)
copy "README_使用说明.txt" "发布版本\" 2>nul || echo [提示] 未找到README文件，跳过...

echo.
echo ========================================
echo ✓ 打包成功！
echo ========================================
echo.
echo 可执行文件位置: dist\迈瑞中国区技术支持数据中心.exe
echo 发布文件夹: 发布版本\
echo.
echo 重要提示：
echo 1. 首次运行需要联网下载Edge浏览器组件
echo 2. 建议在目标机器上先运行一次测试
echo 3. 如需分发，请将"发布版本"文件夹整体拷贝
echo.
pause
