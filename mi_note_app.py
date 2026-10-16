#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
小米笔记导出助手 · 图形界面（MVP-2）
====================================
三个主按钮：① 登录小米账号　② 立即同步（增量）　③ 导出 Word
浏览器抓取在后台线程运行，界面不卡顿；日志与进度实时显示。

运行：python mi_note_app.py（或双击“③打开助手（图形界面）.bat”）
"""
import os
import sys
import queue
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

import mi_note_core as core

# 打包成无控制台 exe 后，sys.stdout/stderr 可能为 None；Playwright 等库若写
# 标准输出会崩溃，这里统一重定向到空设备，保证后台线程稳定。
if getattr(sys, "frozen", False):
    try:
        if sys.stdout is None:
            sys.stdout = open(os.devnull, "w", encoding="utf-8", errors="ignore")
        if sys.stderr is None:
            sys.stderr = open(os.devnull, "w", encoding="utf-8", errors="ignore")
    except Exception:
        pass


CRASH_LOG = core.HERE / "crash.log"


def write_crash(where, exc):
    """把未处理异常写入程序目录下的 crash.log，方便排查（不会自动上传）。"""
    import datetime
    try:
        with open(CRASH_LOG, "a", encoding="utf-8") as f:
            f.write("\n==== %s @ %s ====\n" % (
                where, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            f.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
    except Exception:
        pass


def detect_browser():
    """启动自检：返回常见路径下找到的浏览器，没有则返回 None。"""
    if sys.platform.startswith("win"):
        candidates = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
    return None


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("小米笔记导出助手")
        self.geometry("760x600")
        self.minsize(680, 540)
        self.q = queue.Queue()
        self.worker = None
        self.busy = False

        try:
            ttk.Style(self).theme_use("clam")
        except Exception:
            pass

        self._build_ui()
        self.report_callback_exception = self._on_tk_error
        self._refresh_status()
        self._self_check()
        self.after(150, self._drain)

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}
        font_btn = ("微软雅黑", 11)
        font_main = ("微软雅黑", 10)

        # 标题
        head = ttk.Frame(self)
        head.pack(fill="x", **pad)
        ttk.Label(head, text="小米笔记导出助手",
                  font=("微软雅黑", 16, "bold")).pack(side="left")
        self.status_var = tk.StringVar(value="正在读取本地数据…")
        ttk.Label(head, textvariable=self.status_var,
                  font=("微软雅黑", 9), foreground="#666").pack(side="right")

        # 按钮区
        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=12, pady=(2, 6))
        self.btn_login = ttk.Button(btns, text="① 登录小米账号", command=self.on_login)
        self.btn_sync = ttk.Button(btns, text="② 立即同步（增量）", command=self.on_sync)
        self.btn_export = ttk.Button(btns, text="③ 导出 Word", command=self.on_export)
        self.btn_open = ttk.Button(btns, text="打开文件夹", command=self.on_open_dir)
        self.buttons = [self.btn_login, self.btn_sync, self.btn_export, self.btn_open]
        for b in self.buttons:
            b.pack(side="left", expand=True, fill="x", padx=4)

        # 选项
        opt = ttk.Frame(self)
        opt.pack(fill="x", padx=12)
        self.headless_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="无界面浏览器（已登录后可勾选，后台静默同步）",
                        variable=self.headless_var).pack(side="left")

        # 进度
        prog = ttk.Frame(self)
        prog.pack(fill="x", padx=12, pady=(8, 2))
        self.progress = ttk.Progressbar(prog, mode="determinate", maximum=100)
        self.progress.pack(side="left", fill="x", expand=True)
        self.prog_var = tk.StringVar(value="就绪")
        ttk.Label(prog, textvariable=self.prog_var, width=22,
                  font=("微软雅黑", 9)).pack(side="right", padx=(8, 0))

        # 日志
        logframe = ttk.LabelFrame(self, text="运行日志")
        logframe.pack(fill="both", expand=True, padx=12, pady=8)
        self.logbox = scrolledtext.ScrolledText(
            logframe, wrap="word", font=("Consolas", 10),
            bg="#1e1e1e", fg="#e6e6e6", insertbackground="#e6e6e6",
            relief="flat", state="disabled")
        self.logbox.pack(fill="both", expand=True, padx=6, pady=6)

        # 底部提示
        tip = ("非官方工具 · 仅用于备份你本人的笔记　|　首次使用：先点“① 登录”扫码，"
               "再点“② 立即同步”，完成后点“③ 导出 Word”。登录态会保留几天，日常只需点②③。")
        ttk.Label(self, text=tip, font=("微软雅黑", 8),
                  foreground="#888", wraplength=720, justify="left").pack(
            side="bottom", fill="x", padx=12, pady=(0, 8))

    # ------------------------------------------------------------------ 日志/状态
    def log(self, msg):
        def _do():
            self.logbox.configure(state="normal")
            self.logbox.insert("end", str(msg) + "\n")
            self.logbox.see("end")
            self.logbox.configure(state="disabled")
        self.after(0, _do)

    def set_progress(self, cur, total, text=None):
        def _do():
            if total and total > 0:
                self.progress.configure(mode="determinate", maximum=total, value=cur)
            else:
                self.progress.configure(mode="indeterminate")
            if text:
                self.prog_var.set(text)
        self.after(0, _do)

    def _refresh_status(self):
        try:
            c = core.status()
            self.status_var.set(
                f"本地已有 {c['notes']} 篇 · {c['media']} 个媒体 · {c['folders']} 个分组"
                + (f" · 上次同步 {c['last_sync']}" if c['last_sync'] else " · 尚未同步"))
        except Exception:
            self.status_var.set("尚未同步")

    def _set_busy(self, busy):
        self.busy = busy
        state = "disabled" if busy else "normal"
        for b in self.buttons[:3]:
            b.configure(state=state)
        self.btn_open.configure(state="normal")

    def _on_tk_error(self, exc, val, tb):
        try:
            write_crash("界面回调", val)
            self.log("发生异常，已记录到 crash.log：" + str(val))
        except Exception:
            pass

    def _self_check(self):
        """首次启动环境自检，结果写入日志，降低小白门槛。"""
        self.log("=" * 50)
        self.log("环境自检 ...")
        try:
            import playwright  # noqa: F401
            self.log("  [√] 浏览器自动化组件 Playwright 正常")
        except Exception as e:
            self.log("  [×] 缺少 Playwright，请先双击“①安装依赖.bat”，或在命令行执行：")
            self.log("      py -m pip install -r requirements.txt")
            self.log("      （" + str(e) + "）")
        b = detect_browser()
        if b:
            self.log("  [√] 检测到浏览器：" + os.path.basename(b))
        else:
            self.log("  [!] 常见路径未发现 Edge/Chrome；登录时会尝试内置浏览器，")
            self.log("      若失败请安装 Edge，或运行：py -m playwright install chromium")
        self.log("  数据与导出目录：" + str(core.HERE))
        try:
            c = core.status()
            if c["notes"] == 0:
                self.log("  首次使用：请点“① 登录小米账号”扫码，再点“② 立即同步”。")
            else:
                self.log("  本地已有 %d 篇笔记，可直接点“② 立即同步”做增量。" % c["notes"])
        except Exception:
            pass
        self.log("=" * 50)

    # ------------------------------------------------------------------ 后台线程
    def _start(self, kind, target):
        if self.busy:
            return
        self._set_busy(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)

        def log_cb(m):
            self.q.put(("log", m))

        def prog_cb(phase, cur, total, extra=None):
            extra = extra or {}
            if phase == "notes":
                txt = f"正文 {cur}/{total}（已存 {extra.get('received', 0)}）"
                self.q.put(("progress", cur, total, txt))
            elif phase == "plan":
                self.q.put(("log", f"增量计划完成，待抓取 {total} 篇。"))
            else:
                label = "正在生成 Word…" if phase == "word" else str(phase)
                self.q.put(("progress", cur, total, label))

        def run():
            try:
                target(log_cb, prog_cb)
            except Exception as e:
                write_crash("后台任务:" + kind, e)
                log_cb("出错：" + str(e))
                log_cb("（详细错误已写入程序目录下的 crash.log）")
                self.q.put(("done", kind, False, str(e)))

        threading.Thread(target=run, daemon=True).start()

    # ------------------------------------------------------------------ 按钮动作
    def on_login(self):
        def target(log, prog):
            headless = self.headless_var.get()
            ok = core.login_only(headless=headless, log=log)
            self.q.put(("done", "login", ok, None))
        self.log("=" * 50)
        self._start("login", target)

    def on_sync(self):
        def target(log, prog):
            headless = self.headless_var.get()
            res = core.sync(headless=headless, log=log, progress=prog)
            self.q.put(("done", "sync", bool(res.get("ok")), res))
        self.log("=" * 50)
        self._start("sync", target)

    def on_export(self):
        def target(log, prog):
            path = core.export_word(log=log, progress=prog)
            self.q.put(("done", "export", True, str(path)))
        self.log("=" * 50)
        self._start("export", target)

    def on_open_dir(self):
        try:
            p = str(core.HERE)
            if sys.platform.startswith("win"):
                os.startfile(p)  # noqa: B606
            elif sys.platform == "darwin":
                os.system(f'open "{p}"')
            else:
                os.system(f'xdg-open "{p}"')
        except Exception as e:
            self.log("无法打开文件夹：" + str(e))

    # ------------------------------------------------------------------ 队列轮询
    def _drain(self):
        try:
            while True:
                kind, *rest = self.q.get_nowait()
                if kind == "log":
                    self.log(rest[0])
                elif kind == "progress":
                    cur, total, txt = rest
                    self.set_progress(cur, total, txt)
                elif kind == "done":
                    self._on_done(rest[0], rest[1], rest[2])
        except queue.Empty:
            pass
        self.after(150, self._drain)

    def _on_done(self, kind, ok, payload):
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)
        self._set_busy(False)
        self._refresh_status()
        if kind == "login":
            self.prog_var.set("登录成功" if ok else "登录失败")
            if ok:
                messagebox.showinfo("登录", "登录成功，登录态已保存。可以点“② 立即同步”。")
        elif kind == "sync":
            if ok:
                fetched = payload.get("fetched", 0) if isinstance(payload, dict) else 0
                self.prog_var.set("同步完成")
                self.log(f"本次新抓 {fetched} 篇，本地库已更新。")
                messagebox.showinfo("同步完成",
                                    f"同步完成！本次新抓/更新 {fetched} 篇。\n"
                                    "现在可以点“③ 导出 Word”。")
            else:
                self.prog_var.set("同步未完成")
                messagebox.showwarning("同步", "同步未完成，请查看日志。")
        elif kind == "export":
            if ok:
                self.prog_var.set("Word 已生成")
                messagebox.showinfo("完成", "Word 已生成：\n" + str(payload))
                self.log("Word 已生成：" + str(payload))
            else:
                self.prog_var.set("导出失败")
                messagebox.showerror("导出失败", str(payload))


def main():
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        write_crash("启动", e)
        try:
            messagebox.showerror(
                "启动失败",
                "程序启动失败，错误已写入 crash.log：\n" + str(e))
        except Exception:
            pass


if __name__ == "__main__":
    main()
