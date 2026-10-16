@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==========================================
echo   小米笔记助手 - 打包成免安装 exe（本机做一次）
echo ==========================================
echo.
echo [1/2] 检查/安装 PyInstaller 与依赖 ...
py -m pip install --upgrade pyinstaller
py -m pip install -r requirements.txt
echo.
echo [2/2] 开始打包，约 1-3 分钟，请勿关闭窗口 ...
echo 生成应用图标（若已存在则自动跳过）...
py tools\make_icon.py
py -m PyInstaller mi_note_app.spec --noconfirm
echo.
if exist "dist\小米笔记助手.exe" (
  echo.
  echo 打包成功！产物：dist\小米笔记助手.exe
  echo 便携版直接使用即可；如需安装包与 SHA256 校验值，请运行 ⑤打包发布.bat。
  echo 若首次运行被杀毒软件拦截，请选择“仍要运行/添加信任”。
) else (
  echo 未找到产物，请把本窗口红色报错截图发我。
)
echo.
pause
