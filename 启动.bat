@echo off
title 智能学生排座位系统
cd /d "%~dp0"
echo 正在启动智能学生排座位系统，请稍候（第一次启动会加载图表组件）...
where python >nul 2>nul
if %errorlevel%==0 (
    python "%~dp0src\main.py" %*
) else (
    py -3 "%~dp0src\main.py" %*
)
if errorlevel 1 (
    echo.
    echo 启动遇到问题。请先执行 pip install -r requirements.txt，
    echo 详细步骤见 docs\使用说明.md，并把上面的提示截图发给技术人员。
    echo.
    pause
)
