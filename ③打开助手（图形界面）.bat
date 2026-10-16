@echo off
chcp 65001 >nul
cd /d "%~dp0"
py mi_note_app.py
if errorlevel 1 (
  echo.
  echo 程序异常退出，若缺少依赖请先双击 ①安装依赖.bat
  pause
)
