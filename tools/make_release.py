#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一键打包发布：
  1) PyInstaller 打单文件 exe（便携版）；
  2) 若装了 Inno Setup 6，再编译出 setup.exe 安装包；
  3) 把产物收集到 release/，并生成 SHA256SUMS.txt（用于发布后校验，替代代码签名）。
运行：py tools/make_release.py
"""
import hashlib
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DIST = ROOT / "dist"
RELEASE = ROOT / "release"
EXE = DIST / "小米笔记助手.exe"
ISS = ROOT / "installer" / "mi-note-assistant.iss"
ICON = ROOT / "assets" / "icon.ico"


def run(cmd, cwd=None):
    print(">>", " ".join(str(c) for c in cmd))
    subprocess.run([str(c) for c in cmd], cwd=str(cwd or ROOT), check=True)


def ensure_icon():
    """图标不入库时（如刚 clone），用 Pillow 现场生成，保证打包不缺图标。"""
    if ICON.exists():
        return
    print("未发现 assets/icon.ico，先运行 tools/make_icon.py 生成图标 ...")
    run([sys.executable, str(ROOT / "tools" / "make_icon.py")])


def find_iscc():
    for p in (
        r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        r"C:\Program Files\Inno Setup 6\ISCC.exe",
    ):
        if Path(p).exists():
            return Path(p)
    hits = list(Path(r"C:\Program Files (x86)").glob("Inno Setup*/ISCC.exe"))
    return hits[0] if hits else None


def sha256_of(path: Path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def main():
    # 1) 依赖自检
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("未安装 PyInstaller，正在安装 ...")
        run([sys.executable, "-m", "pip", "install", "pyinstaller"])

    # 2) 确保图标存在 + 打 exe
    ensure_icon()
    run([sys.executable, "-m", "PyInstaller", "mi_note_app.spec", "--noconfirm"])
    if not EXE.exists():
        raise SystemExit("未找到打包产物：%s" % EXE)

    # 3) 收集到 release/
    RELEASE.mkdir(exist_ok=True)
    for old in RELEASE.glob("*"):
        if old.is_file():
            old.unlink()
    shutil.copy2(EXE, RELEASE / EXE.name)

    # 4) 可选：安装包
    iscc = find_iscc()
    if iscc:
        run([iscc, str(ISS)])
        out_dir = ROOT / "installer" / "Output"
        for setup in out_dir.glob("*.exe"):
            shutil.copy2(setup, RELEASE / setup.name)
    else:
        print("[提示] 未安装 Inno Setup 6，跳过安装包（只生成便携 exe）。")
        print("       如需开始菜单/卸载器的安装包，请到 https://jrsoftware.org/isdl.php 安装。")

    # 5) 生成 SHA256
    lines = []
    for f in sorted(RELEASE.glob("*.exe")):
        digest = sha256_of(f)
        lines.append("%s  %s" % (digest, f.name))
    sums = RELEASE / "SHA256SUMS.txt"
    sums.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print("\n发布产物已就绪：%s" % RELEASE)
    for line in lines:
        print("  " + line)
    print("\n校验方法：certutil -hashfile 文件名 SHA256")


if __name__ == "__main__":
    main()
