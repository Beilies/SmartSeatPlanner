@echo off
title 智能学生排座位系统 · 自检
cd /d "%~dp0.."
set "NO_PAUSE=%~1"
echo [1/2] 全链路自检（脏 Excel → 导入 → 5 套方案 → 三种导出）...
python src\selfcheck.py
if errorlevel 1 goto fail
echo.
echo [2/2] 回归测试（需要 pytest：pip install pytest）...
python -m pytest -q tests
if errorlevel 1 goto fail
echo.
echo [完成] 全部通过。
if /i not "%NO_PAUSE%"=="--no-pause" pause
exit /b 0
:fail
echo.
echo [失败] 请查看上面的错误信息。
if /i not "%NO_PAUSE%"=="--no-pause" pause
exit /b 1
