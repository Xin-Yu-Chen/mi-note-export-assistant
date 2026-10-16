# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置：把图形助手打成单个 exe。
# 用法（Windows）：pyinstaller mi_note_app.spec --noconfirm
# 想要“一个文件夹”而非单 exe（启动更快、杀软误报更少），见文件末尾注释。

from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []

# Playwright 自带 node driver 与大量 JS 资源，必须整包收集；
# docx / PIL 一并收集，避免子模块漏装。
for pkg in ("playwright", "docx", "PIL"):
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

# 浏览器使用系统自带的 Edge（channel=msedge），因此不打包 chromium，体积小很多。

# 图标可由 tools/make_icon.py 现场生成；若缺失则不使用自定义图标（不阻断打包）
import os as _os
_icon_path = _os.path.join(SPECPATH, "assets", "icon.ico")
icon_path = _icon_path if _os.path.exists(_icon_path) else None

a = Analysis(
    ["mi_note_app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "numpy", "pandas", "matplotlib", "scipy", "pytest",
        "PyQt5", "PyQt6", "PySide2", "PySide6", "IPython", "notebook",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

# 单文件 exe（onefile）
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="小米笔记助手",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # 不弹黑色控制台窗口
    disable_windowed_traceback=False,
    icon=icon_path,
    version=r"assets/version_info.txt",
)

# ---------------------------------------------------------------------------
# 若 onefile 被杀软误报或启动慢，改用“文件夹版（onedir）”：
#   1) 把上面 EXE(...) 里的 a.binaries、a.datas 两个参数删掉；
#   2) 在文件末尾追加：
#        coll = COLLECT(
#            EXE(pyz, a.scripts, [], name="小米笔记助手", console=False),
#            a.binaries, a.datas, strip=False, upx=True, name="小米笔记助手")
#   3) 重新 pyinstaller，产物在 dist\小米笔记助手\ 文件夹，双击其中的 exe。
# ---------------------------------------------------------------------------
