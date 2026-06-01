@echo off
chcp 65001 >nul
echo ========================================
echo 公众号数据工具 - 构建脚本
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python，请先安装Python 3.8+
    pause
    exit /b 1
)

echo [1/4] 检查Python环境... OK
python --version
echo.

REM 检查并安装PyInstaller
echo [2/4] 检查PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo 正在安装PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo [错误] PyInstaller安装失败
        pause
        exit /b 1
    )
)
echo PyInstaller已就绪
echo.

REM 安装项目依赖
echo [3/4] 安装项目依赖...
pip install -r requirements.txt
if errorlevel 1 (
    echo [警告] 部分依赖安装可能有问题，继续尝试构建...
)
echo.

REM 使用PyInstaller构建
echo [4/4] 开始构建可执行文件...
echo 这可能需要几分钟时间，请耐心等待...
echo.

pyinstaller --noconfirm --clean --onefile --windowed --name gzh_report_tool main.py

if errorlevel 1 (
    echo.
    echo [错误] 构建失败！
    pause
    exit /b 1
)

echo.
echo ========================================
echo 构建成功！
echo ========================================
echo.
echo 可执行文件位置: dist\gzh_report_tool.exe
echo.
echo 你可以:
echo 1. 直接双击 dist\gzh_report_tool.exe 运行
echo 2. 将整个 dist 文件夹复制到其他电脑上使用
echo.
pause
