@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==========================================
echo   小米笔记助手 - 安装/检查依赖（仅首次）
echo ==========================================
echo.
py -m pip install -r requirements.txt
echo.
echo 依赖处理完成。若上方出现红色报错，可截图反馈。
pause
