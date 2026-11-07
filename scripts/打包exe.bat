@echo off
title 打包智能排座位系统 · 免安装 exe
cd /d "%~dp0.."

echo ============================================================
echo   一键打包 SmartSeatPlanner（Windows 单文件 exe）
echo   产物会放在 release\，该目录已由 .gitignore 排除
echo ============================================================
echo.
set "NO_PAUSE=%~1"

where python >nul 2>nul
if errorlevel 1 (
    echo [失败] 没找到 Python，请先安装官方 Python 3.10+ 并勾选 Add Python to PATH。
    if /i not "%NO_PAUSE%"=="--no-pause" pause
    exit /b 1
)

python -c "import PyInstaller" >nul 2>nul
if errorlevel 1 (
    echo [准备] 首次打包需要 PyInstaller，正在安装...
    python -m pip install pyinstaller
)

python -c "import pandas, openpyxl, xlrd, matplotlib, PIL" >nul 2>nul
if errorlevel 1 (
    echo [准备] 缺少运行依赖，正在安装...
    python -m pip install -r requirements.txt
)

python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name SmartSeatPlanner --icon "assets\app.ico" --distpath "release" ^
  --paths "src" ^
  --add-data "assets\app.ico;." --add-data "assets\app_logo.png;." ^
  --hidden-import selfcheck --hidden-import app_paths ^
  --hidden-import matplotlib.backends.backend_tkagg --hidden-import tkinter.ttk ^
  --collect-data matplotlib "src\main.py"

if errorlevel 1 (
    echo.
    echo [失败] 打包没有完成，请把上面的日志发给技术人员。
    if /i not "%NO_PAUSE%"=="--no-pause" pause
    exit /b 1
)

move /y "release\SmartSeatPlanner.exe" "release\智能排座位系统.exe" >nul

echo [整理] 把老师需要的文件一起放进 release\ ...
copy /y "docs\免安装版-先看我.txt" "release\" >nul
copy /y "docs\使用说明.md" "release\" >nul
copy /y "scripts\一键放到桌面.bat" "release\" >nul
copy /y "examples\测试数据_60人成绩表.xlsx" "release\" >nul
copy /y "examples\测试数据_60人_精简版.xlsx" "release\" >nul

echo [验证] 运行 exe --diagnose 检查打包完整性...
"release\智能排座位系统.exe" --diagnose
if errorlevel 1 (
    echo [失败] 自检未通过，请看 release\打包自检结果.txt。
    if /i not "%NO_PAUSE%"=="--no-pause" pause
    exit /b 1
)

echo.
echo [完成] release\智能排座位系统.exe 已就绪。
echo        把整个 release\ 文件夹压缩后上传到 GitHub Releases 即可。
echo        详细步骤见 docs\发布指南.md。
echo.
if /i not "%NO_PAUSE%"=="--no-pause" pause
