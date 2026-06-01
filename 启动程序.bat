@echo off
setlocal EnableExtensions

REM Use UTF-8 for better Chinese output
chcp 65001 >nul

REM Switch to script directory (supports spaces)
cd /d "%~dp0"

echo ============================================
echo 公众号数据整合工具 - 一键运行
echo CurrentDir: "%cd%"
echo ============================================
echo.

REM 1) Check Python
python --version >nul 2>&1
if errorlevel 1 (
  echo [错误] 未检测到 Python，请先安装 Python 3.8+ 并勾选 "Add Python to PATH"。
  echo        安装后重新双击本文件即可。
  echo.
  pause
  exit /b 1
)

REM 2) Upgrade pip (best effort)
python -m pip --version >nul 2>&1
if errorlevel 1 (
  echo [错误] pip 不可用。请尝试重装 Python，或确保已安装 pip。
  echo.
  pause
  exit /b 1
)

echo [1/3] 检查并安装依赖...
if exist "requirements.txt" (
  python -m pip install -r "requirements.txt"
) else (
  python -m pip install pandas openpyxl xlrd
)
if errorlevel 1 (
  echo.
  echo [错误] 依赖安装失败。请检查网络，或稍后重试。
  echo.
  pause
  exit /b 1
)

REM 3) Run script
echo.
echo [2/3] 启动GUI界面...
set PYTHONIOENCODING=utf-8
python "gui.py"
if errorlevel 1 (
  echo.
  echo [错误] 运行失败。请把本窗口报错信息截图发我，我来帮你定位。
  echo.
  pause
  exit /b 1
)

echo.
echo [3/3] 完成。正在打开输出文件夹...
if not exist "output" mkdir "output" >nul 2>&1
explorer "output"

echo.
echo 你可以在 output\ 里找到生成的 Excel 文件。
pause
exit /b 0

