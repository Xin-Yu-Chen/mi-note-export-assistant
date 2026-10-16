@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo  正在一键打包（exe + 可选安装包 + SHA256校验）
echo  约需 1-3 分钟，请勿关闭窗口 ...
echo ============================================
py tools\make_release.py
echo.
echo 完成！发布产物在 release 文件夹（含 SHA256SUMS.txt）。
pause