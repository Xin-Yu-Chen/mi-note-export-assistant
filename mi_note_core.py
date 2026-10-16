#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
小米 / 红米笔记 · 核心逻辑（无界面，供 GUI 与命令行共用）
========================================================
职责：
  1. 本地 SQLite 真源（mi_notes.db）：笔记 / 媒体 / 分组 / 同步状态。
  2. Playwright 登录 + 两阶段增量抓取：
       阶段一：拉全量列表（只含 id/时间，很便宜），与本地比对，算出
               新增 / 更新 / 已删除的笔记；
       阶段二：只对“新增或被修改”的笔记拉全文 + 图片，其余直接跳过。
     首次同步本地为空 -> 全量；以后每天只拉变化的几篇。
  3. 导出 Word：优先调用同目录的 export_mi_notes.py（沿用你已有的排版），
     找不到时用内置的基础渲染兜底。

全程只读云端、只连官方 i.mi.com；登录态存于 .browser-profile。
"""
import os
import re
import io
import sys
import json
import time
import base64
import sqlite3
import platform
import subprocess
import threading
from pathlib import Path

# ---------------------------------------------------------------------------
# 路径与常量
# ---------------------------------------------------------------------------
APP_DIR_NAME = "MiNoteExportAssistant"


def _writable(d: Path) -> bool:
    """判断目录是否可写（用于判断 exe 装在 Program Files 这类只读位置时的数据落点）。"""
    try:
        d.mkdir(parents=True, exist_ok=True)
        probe = d / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return True
    except Exception:
        return False


def _app_dir() -> Path:
    """数据/配置目录（数据库、登录态、导出的 Word 都在这里）。
    - 普通脚本运行：本文件所在目录；
    - exe 版：优先 exe 所在目录（便携模式，数据随程序走）；
      若 exe 所在目录不可写（如被安装到 Program Files），
      则改用 %APPDATA%\\MiNoteExportAssistant（用户可写目录）。
    """
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if _writable(exe_dir):
            return exe_dir
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        user_dir = Path(base) / APP_DIR_NAME
        user_dir.mkdir(parents=True, exist_ok=True)
        return user_dir
    return Path(__file__).resolve().parent


HERE = _app_dir()
FROZEN = bool(getattr(sys, "frozen", False))
DB_PATH = HERE / "mi_notes.db"
PROFILE_DIR = HERE / ".browser-profile"
EXPORT_JSON = HERE / "mi_notes_export.json"
WORD_EXPORTER = HERE / "export_mi_notes.py"

NOTE_URL = "https://i.mi.com/note/h5"
LOGIN_WAIT_SECONDS = 900
PULL_BATCH = 20
PULL_INTERVAL_MS = 2000
LIST_WAIT_SECONDS = 120

# ---------------------------------------------------------------------------
# 页面脚本
# ---------------------------------------------------------------------------

# 阶段一：启动“拉全量列表”的后台任务（立即返回，不阻塞）
LIST_JS = r"""
() => {
  const BASE = "https://i.mi.com";
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const now = () => Date.now();
  const S = window.__MI = {
    phase: "list", listReady: false, _listSent: false,
    folders: {}, entries: [],
    pending: {}, nextSeq: 0,
    fetchStarted: false, total: 0, i: 0, kept: 0, mediaCount: 0,
    done: false, error: null
  };
  const getJSON = async (url) => {
    const r = await fetch(url, { credentials: "include", headers: { "Accept": "application/json" } });
    if (r.status === 401) throw new Error("未登录(401)");
    if (!r.ok) throw new Error("HTTP " + r.status + " " + url);
    return r.json();
  };
  window.__MI_list = (async () => {
    try {
      let syncTag = "", entries = [], page = 0;
      while (true) {
        page++;
        let url = `${BASE}/note/full/page/?ts=${now()}&limit=200`;
        if (syncTag) url += `&syncTag=${syncTag}`;
        const j = await getJSON(url);
        if (j.result !== "ok" || !j.data) throw new Error("列表接口异常");
        const d = j.data;
        (d.folders || []).forEach((f) => { S.folders[f.id] = f.subject; });
        entries = entries.concat(d.entries || []);
        S.entries = entries;
        syncTag = d.syncTag || syncTag;
        if (d.lastPage || !d.syncTag) break;
        await sleep(250);
      }
      S.phase = "await_plan";
      S.listReady = true;
    } catch (err) {
      S.error = String((err && err.message) || err);
      S.phase = "error";
      S.done = true;
    }
  })();
  return { started: true };
}
"""

# 取列表（只在就绪后成功返回一次精简清单）
GET_LIST_JS = r"""
() => {
  const S = window.__MI;
  if (!S) return null;
  const base = {
    listReady: S.listReady,
    status: { phase: S.phase, total: S.entries.length, done: S.done, error: S.error }
  };
  if (!S.listReady) return base;
  if (S._listSent) return Object.assign(base, { resent: true });
  S._listSent = true;
  return Object.assign(base, {
    folders: S.folders,
    entries: S.entries.map((e) => ({
      id: String(e.id),
      modifyDate: e.modifyDate || 0,
      createDate: e.createDate || 0,
      folderId: String(e.folderId || "")
    }))
  });
}
"""

# 阶段二：传入“需要拉全文”的 id 列表，仅抓这些（抓完放进 pending）
START_FETCH_JS = r"""
(ids) => {
  const BASE = "https://i.mi.com";
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const now = () => Date.now();
  const S = window.__MI;
  if (!S || !S.listReady) return { ok: false, reason: "list_not_ready" };
  if (S.fetchStarted) return { ok: false, reason: "already" };
  S.fetchStarted = true;
  S.total = ids.length;
  S.phase = ids.length ? "notes" : "done";
  const byId = {};
  S.entries.forEach((e) => { byId[String(e.id)] = e; });

  const getJSON = async (url) => {
    const r = await fetch(url, { credentials: "include", headers: { "Accept": "application/json" } });
    if (r.status === 401) throw new Error("未登录(401)");
    if (!r.ok) throw new Error("HTTP " + r.status + " " + url);
    return r.json();
  };
  const parseExtra = (e) => { try { return JSON.parse(e.extraInfo || "{}"); } catch (_) { return {}; } };
  const getFiles = (e) => {
    let s = e.setting;
    if (typeof s === "string") { try { s = JSON.parse(s); } catch (_) { s = null; } }
    const arr = (s && s.data) || e.files || [];
    return Array.isArray(arr) ? arr : [];
  };
  const blobToDataURL = (blob) => new Promise((res, rej) => {
    const fr = new FileReader(); fr.onload = () => res(fr.result); fr.onerror = rej;
    fr.readAsDataURL(blob);
  });

  if (!ids.length) { S.done = true; return { ok: true, count: 0 }; }

  window.__MI_fetch = (async () => {
    try {
      let kept = 0;
      for (let n = 0; n < ids.length; n++) {
        S.i = n + 1;
        const e0 = byId[ids[n]];
        if (!e0) { await sleep(20); continue; }
        try {
          const jd = await getJSON(`${BASE}/note/note/${e0.id}/?ts=${now()}`);
          const e = jd && jd.data && jd.data.entry ? Object.assign({}, e0, jd.data.entry) : e0;
          const extra = parseExtra(e);
          if (extra.note_content_type === "mind" || extra.mind_content) { await sleep(30); continue; }
          const content = e.content || e.snippet || "";
          if (!content) { await sleep(30); continue; }
          const files = getFiles(e)
            .map((f) => ({ rawId: f.rawId || f.fileId || f.digest, mimeType: f.mimeType || "" }))
            .filter((f) => f.rawId);
          const mediaItems = [];
          const mediaRe = /<(?:img|sound|video)\b[^>]*?(?:id|data)=["']([^"']*)["'][^>]*?\/?>/g;
          let mm; const idset = new Set();
          while ((mm = mediaRe.exec(content))) idset.add(mm[1]);
          for (const rawId of idset) {
            const meta = files.find((f) => f.rawId === rawId) || {};
            const fileId = rawId.indexOf(".") >= 0 ? rawId.split(".")[1] : rawId;
            try {
              const rr = await fetch(
                `${BASE}/file/full?type=note_img&fileid=${encodeURIComponent(fileId)}&ts=${now()}`,
                { credentials: "include" });
              if (rr.ok) {
                const dataURL = await blobToDataURL(await rr.blob());
                mediaItems.push({ rawId, mime: meta.mimeType || rr.headers.get("content-type") || "", data: dataURL });
              }
            } catch (_) {}
            await sleep(30);
          }
          S.mediaCount += mediaItems.length;
          S.pending[S.nextSeq] = {
            id: String(e.id), folderId: String(e.folderId || ""),
            title: (extra.title || e.subject || "").trim(),
            createDate: e.createDate || 0, modifyDate: e.modifyDate || 0,
            content, files, mediaItems
          };
          S.nextSeq++;
          S.kept = ++kept;
          await sleep(110);
        } catch (err) { await sleep(100); }
      }
      S.phase = "done";
      S.done = true;
    } catch (err) {
      S.error = String((err && err.message) || err);
      S.phase = "error";
      S.done = true;
    }
  })();
  return { ok: true, count: ids.length };
}
"""

# 分批取走已抓取的笔记（取走即从页面内存删除）
PULL_JS = r"""
(start) => {
  const S = window.__MI;
  if (!S) return null;
  const notes = [];
  let k = start;
  while (S.pending[k] !== undefined && notes.length < %d) {
    notes.push(S.pending[k]);
    delete S.pending[k];
    k++;
  }
  return {
    ackNext: k,
    status: {
      phase: S.phase, total: S.total, i: S.i, kept: S.kept,
      mediaCount: S.mediaCount, done: S.done, error: S.error
    },
    notes
  };
}
""" % PULL_BATCH

# 探测登录态
CHECK_LOGIN_JS = r"""
async () => {
  try {
    const r = await fetch('https://i.mi.com/note/full/page/?ts=' + Date.now() + '&limit=1',
                          { credentials: 'include', headers: { 'Accept': 'application/json' } });
    if (r.status === 401) return false;
    const j = await r.json();
    return j.result === 'ok';
  } catch (e) { return false; }
}
"""


def _noop(*_a, **_kw):
    pass


# ---------------------------------------------------------------------------
# 本地数据库（SQLite 真源）
# ---------------------------------------------------------------------------
class NotesDB:
    def __init__(self, path: Path = DB_PATH):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.RLock()
        self.conn = sqlite3.connect(str(self.path), check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._init_schema()

    def _init_schema(self):
        with self._lock, self.conn:
            self.conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS notes (
                    id TEXT PRIMARY KEY,
                    folder_id TEXT,
                    title TEXT,
                    create_date INTEGER,
                    modify_date INTEGER,
                    content TEXT,
                    files TEXT,
                    deleted INTEGER DEFAULT 0,
                    synced_at INTEGER
                );
                CREATE TABLE IF NOT EXISTS media (
                    raw_id TEXT PRIMARY KEY,
                    note_id TEXT,
                    mime TEXT,
                    data TEXT
                );
                CREATE TABLE IF NOT EXISTS folders (
                    id TEXT PRIMARY KEY,
                    name TEXT
                );
                CREATE TABLE IF NOT EXISTS meta (
                    key TEXT PRIMARY KEY,
                    value TEXT
                );
                CREATE INDEX IF NOT EXISTS idx_notes_mod ON notes(modify_date);
                CREATE INDEX IF NOT EXISTS idx_notes_folder ON notes(folder_id);
                """
            )

    # ---- meta ----
    def get_meta(self, key, default=None):
        with self._lock:
            row = self.conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default

    def set_meta(self, key, value):
        with self._lock, self.conn:
            self.conn.execute(
                "INSERT INTO meta(key,value) VALUES(?,?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, str(value)))

    def get_last_sync(self):
        return self.get_meta("last_sync", "")

    # ---- folders ----
    def upsert_folders(self, folders: dict):
        if not folders:
            return
        with self._lock, self.conn:
            for fid, name in folders.items():
                self.conn.execute(
                    "INSERT INTO folders(id,name) VALUES(?,?) "
                    "ON CONFLICT(id) DO UPDATE SET name=excluded.name",
                    (str(fid), name or ""))

    # ---- 同步计划：对比全量列表，决定哪些要拉全文、哪些云端已删 ----
    def plan(self, entries: list):
        """返回 (to_fetch, cloud_ids, stats)。to_fetch 为需要拉全文的 id 列表。"""
        now_ms = int(time.time() * 1000)
        cloud_ids = {e["id"] for e in entries}
        cloud_mod = {e["id"]: int(e.get("modifyDate") or 0) for e in entries}
        with self._lock:
            local = {r["id"]: r for r in
                     self.conn.execute("SELECT id, modify_date, deleted FROM notes").fetchall()}

        new_ids, updated_ids, restored_ids, deleted_ids = [], [], [], []
        for cid in cloud_ids:
            row = local.get(cid)
            if row is None:
                new_ids.append(cid)
            elif row["deleted"]:
                # 云端重新出现（之前被删）-> 复位并重新拉全文
                restored_ids.append(cid)
            elif cloud_mod[cid] > int(row["modify_date"] or 0):
                updated_ids.append(cid)
        for lid, row in local.items():
            if not row["deleted"] and lid not in cloud_ids:
                deleted_ids.append(lid)

        # 复位“恢复”的笔记
        if restored_ids:
            with self._lock, self.conn:
                self.conn.executemany(
                    "UPDATE notes SET deleted=0 WHERE id=?", [(i,) for i in restored_ids])

        to_fetch = new_ids + updated_ids + restored_ids
        stats = {
            "new": len(new_ids), "updated": len(updated_ids),
            "restored": len(restored_ids), "deleted": len(deleted_ids),
            "unchanged": len(cloud_ids) - len(new_ids) - len(updated_ids)
                         - len(restored_ids),
            "total_cloud": len(cloud_ids), "to_fetch": len(to_fetch),
        }
        return to_fetch, deleted_ids, stats

    # ---- 写笔记 ----
    def upsert_note(self, note: dict):
        with self._lock, self.conn:
            self.conn.execute(
                """
                INSERT INTO notes(id,folder_id,title,create_date,modify_date,content,files,deleted,synced_at)
                VALUES(?,?,?,?,?,?,?,0,?)
                ON CONFLICT(id) DO UPDATE SET
                    folder_id=excluded.folder_id, title=excluded.title,
                    create_date=excluded.create_date, modify_date=excluded.modify_date,
                    content=excluded.content, files=excluded.files,
                    deleted=0, synced_at=excluded.synced_at
                """,
                (note["id"], note.get("folderId", ""), note.get("title", ""),
                 int(note.get("createDate") or 0), int(note.get("modifyDate") or 0),
                 note.get("content", ""), json.dumps(note.get("files", []), ensure_ascii=False),
                 int(time.time() * 1000)))
            for m in note.get("mediaItems", []) or []:
                self.conn.execute(
                    """
                    INSERT INTO media(raw_id,note_id,mime,data) VALUES(?,?,?,?)
                    ON CONFLICT(raw_id) DO UPDATE SET
                        note_id=excluded.note_id, mime=excluded.mime, data=excluded.data
                    """, (m["rawId"], note["id"], m.get("mime", ""), m.get("data", "")))

    def mark_deleted(self, ids: list):
        if not ids:
            return
        with self._lock, self.conn:
            self.conn.executemany(
                "UPDATE notes SET deleted=1 WHERE id=?", [(i,) for i in ids])

    # ---- 统计 ----
    def counts(self):
        with self._lock:
            notes = self.conn.execute(
                "SELECT COUNT(*) c FROM notes WHERE deleted=0").fetchone()["c"]
            media = self.conn.execute("SELECT COUNT(*) c FROM media").fetchone()["c"]
            folders = self.conn.execute("SELECT COUNT(*) c FROM folders").fetchone()["c"]
        return {"notes": notes, "media": media, "folders": folders,
                "last_sync": self.get_last_sync()}

    # ---- 导出（结构对齐旧的 mi_notes_data.json）----
    def export_json(self, include_deleted=False):
        with self._lock:
            sql = ("SELECT id,folder_id,title,create_date,modify_date,content,files "
                   "FROM notes WHERE deleted=0 ORDER BY create_date ASC, id ASC")
            if include_deleted:
                sql = ("SELECT id,folder_id,title,create_date,modify_date,content,files "
                       "FROM notes ORDER BY create_date ASC, id ASC")
            rows = self.conn.execute(sql).fetchall()
            mrows = self.conn.execute("SELECT raw_id,mime,data FROM media").fetchall()
            frows = self.conn.execute("SELECT id,name FROM folders").fetchall()
        notes = []
        for r in rows:
            try:
                files = json.loads(r["files"] or "[]")
            except Exception:
                files = []
            notes.append({
                "id": r["id"], "folderId": r["folder_id"] or "",
                "title": r["title"] or "",
                "createDate": r["create_date"] or 0,
                "modifyDate": r["modify_date"] or 0,
                "content": r["content"] or "", "files": files,
            })
        media = {r["raw_id"]: {"mime": r["mime"] or "", "data": r["data"] or ""}
                 for r in mrows}
        folders = {r["id"]: r["name"] or "" for r in frows}
        return {
            "exportedAt": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "folders": folders, "notes": notes, "media": media,
        }

    def close(self):
        with self._lock:
            self.conn.close()


# ---------------------------------------------------------------------------
# 浏览器
# ---------------------------------------------------------------------------
def launch_context(p, headless: bool):
    common = dict(
        user_data_dir=str(PROFILE_DIR),
        headless=headless,
        accept_downloads=True,
        viewport={"width": 1280, "height": 860},
        args=["--no-sandbox", "--disable-dev-shm-usage",
              "--disable-blink-features=AutomationControlled"],
    )
    if platform.system() == "Windows":
        for channel in ("msedge", "chrome"):
            try:
                return p.chromium.launch_persistent_context(channel=channel, **common)
            except Exception:
                continue
        try:
            return p.chromium.launch_persistent_context(**common)
        except Exception as e:
            raise RuntimeError(
                "未找到可用浏览器（系统 Edge / Chrome 都没有，且内置 Chromium 未安装）。\n"
                "请任选其一解决后重试：\n"
                "  1) 安装 Microsoft Edge（Windows 一般自带，推荐）；\n"
                "  2) 或在本目录命令行执行：  py -m playwright install chromium\n"
                "（原始信息：%s）" % e)
    for exe in ("/usr/local/bin/chromium", "/usr/bin/chromium",
                "/usr/bin/chromium-browser", "/usr/bin/google-chrome"):
        if Path(exe).exists():
            return p.chromium.launch_persistent_context(executable_path=exe, **common)
    try:
        return p.chromium.launch_persistent_context(**common)
    except Exception as e:
        raise RuntimeError(
            "未找到可用浏览器。请安装 Chromium/Chrome，或执行："
            " python -m playwright install chromium（原始信息：%s）" % e)


def _is_logged_in(page):
    for _ in range(3):
        try:
            return bool(page.evaluate(CHECK_LOGIN_JS))
        except Exception:
            page.wait_for_timeout(1500)
    return False


def _open_and_wait_login(page, headless, log):
    """打开笔记页，必要时等待用户手动登录。返回是否已登录。"""
    log("打开小米云笔记页面 ...")
    page.goto(NOTE_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(4000)
    if _is_logged_in(page):
        log("✅ 已存在有效登录态，免登录。")
        return True
    if headless:
        log("未登录且为无界面模式，无法登录。请先在界面模式下登录一次。")
        return False
    log("")
    log("=" * 56)
    log("浏览器已打开，请在窗口里登录小米账号（可扫码）。")
    log("登录成功并能看到笔记后，程序会自动继续。")
    log("=" * 56)
    deadline = time.time() + LOGIN_WAIT_SECONDS
    while time.time() < deadline:
        page.wait_for_timeout(3000)
        if _is_logged_in(page):
            log("✅ 检测到登录成功！")
            if "/note" not in (page.url or ""):
                page.goto(NOTE_URL, wait_until="domcontentloaded")
                page.wait_for_timeout(3000)
            return True
    log("等待登录超时。")
    return False


LEGACY_JSON = HERE / "mi_notes_data.json"


def import_legacy_json(db: "NotesDB", path: Path, log=_noop):
    """把 MVP-1 全量导出的 mi_notes_data.json 导入 SQLite，避免首次重新全量抓取。"""
    path = Path(path)
    with open(path, "r", encoding="utf-8") as f:
        doc = json.load(f)
    db.upsert_folders(doc.get("folders", {}))
    media = doc.get("media", {})
    n = 0
    for note in doc.get("notes", []):
        items = []
        for fmeta in note.get("files", []):
            rid = fmeta.get("rawId")
            if rid in media:
                items.append({"rawId": rid,
                              "mime": media[rid].get("mime", ""),
                              "data": media[rid].get("data", "")})
        note = dict(note)
        note["mediaItems"] = items
        db.upsert_note(note)
        n += 1
    db.set_meta("last_sync", doc.get("exportedAt", "") or
                time.strftime("%Y-%m-%d %H:%M:%S"))
    log(f"已从历史文件 {path.name} 导入 {n} 篇笔记，接下来只做增量。")
    return n


def login_only(headless=False, log=_noop):
    """仅建立/刷新登录态，登录成功后关闭浏览器。"""
    from playwright.sync_api import sync_playwright
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        log("正在启动浏览器 ...")
        context = launch_context(p, headless)
        page = context.pages[0] if context.pages else context.new_page()
        ok = _open_and_wait_login(page, headless, log)
        if ok:
            log("登录态已保存，之后同步一般无需再次登录。")
        try:
            context.close()
        except Exception:
            pass
        return ok


def sync(headless=False, log=_noop, progress=_noop, stop_check=None,
         db_path: Path = DB_PATH):
    """
    执行一次增量同步。
    progress(phase, current, total, extra_dict)
    返回统计 dict；stop_check() 返回 True 时尽量提前收尾。
    """
    from playwright.sync_api import sync_playwright
    stop_check = stop_check or (lambda: False)
    db = NotesDB(db_path)
    result = {"ok": False}
    # 本地库为空时，先导入历史全量 JSON，避免首次再全量抓取一遍
    if db.counts()["notes"] == 0 and LEGACY_JSON.exists():
        try:
            import_legacy_json(db, LEGACY_JSON, log)
        except Exception as e:
            log("历史数据导入失败（忽略，将全量抓取）：%s" % e)
    try:
        PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        with sync_playwright() as p:
            log("正在启动浏览器 ...")
            context = launch_context(p, headless)
            page = context.pages[0] if context.pages else context.new_page()

            if not _open_and_wait_login(page, headless, log):
                context.close()
                result["error"] = "未登录"
                return result

            # 阶段一：拉全量列表
            log("正在获取笔记列表（不含正文，很快）...")
            page.evaluate(LIST_JS)
            payload = None
            t0 = time.time()
            while time.time() - t0 < LIST_WAIT_SECONDS:
                g = page.evaluate(GET_LIST_JS)
                if g and g.get("listReady") and "entries" in g:
                    payload = g
                    break
                if g and g.get("status", {}).get("error"):
                    raise RuntimeError(g["status"]["error"])
                page.wait_for_timeout(1500)
            if payload is None:
                raise RuntimeError("获取列表超时")

            entries = payload["entries"]
            folders = payload.get("folders", {})
            db.upsert_folders(folders)
            to_fetch, deleted_ids, stats = db.plan(entries)
            log(f"云端共 {stats['total_cloud']} 篇：新增 {stats['new']}、"
                f"更新 {stats['updated'] + stats['restored']}、"
                f"未变化 {stats['unchanged']}、云端已删除 {stats['deleted']}。")
            if deleted_ids:
                db.mark_deleted(deleted_ids)
            progress("plan", 0, max(1, len(to_fetch)), stats)

            # 阶段二：只拉变化的笔记
            if not to_fetch:
                log("✅ 没有需要更新的笔记，已是最新。")
                db.set_meta("last_sync", time.strftime("%Y-%m-%d %H:%M:%S"))
                result.update(ok=True, fetched=0, **stats)
                try:
                    context.close()
                except Exception:
                    pass
                return result

            log(f"开始抓取 {len(to_fetch)} 篇新增/修改笔记的正文 ...")
            started = page.evaluate(START_FETCH_JS, to_fetch)
            if not started or not started.get("ok"):
                raise RuntimeError("无法启动正文抓取：%s" % started)

            ack = 0
            received = 0
            last_save_tip = 0.0
            deadline = time.time() + max(600, len(to_fetch) * 8)
            fatal = None
            while True:
                if stop_check():
                    log("收到停止指令，准备收尾（已抓内容已保存）。")
                    break
                r = None
                try:
                    r = page.evaluate(PULL_JS, ack)
                except Exception:
                    page.wait_for_timeout(2000)
                    try:
                        r = page.evaluate(PULL_JS, ack)
                    except Exception:
                        r = None
                if r:
                    for n in r.get("notes", []):
                        db.upsert_note(n)
                        received += 1
                    ack = r.get("ackNext", ack)
                    st = r.get("status", {})
                    progress("notes", st.get("i", 0), st.get("total", len(to_fetch)),
                             {"received": received, "media": st.get("mediaCount", 0)})
                    if time.time() - last_save_tip > 3:
                        log(f"  正文 {st.get('i', 0)}/{st.get('total', len(to_fetch))}，"
                            f"已保存 {received} 篇，媒体 {st.get('mediaCount', 0)} 个")
                        last_save_tip = time.time()
                    if st.get("error"):
                        fatal = st["error"]
                    if st.get("done") and received >= st.get("kept", 0) and not r.get("notes"):
                        break
                if fatal:
                    log("页面抓取报错：%s（已抓部分已保存）" % fatal)
                    break
                if time.time() > deadline:
                    log("抓取超时，先保存已抓到的部分。")
                    break
                page.wait_for_timeout(PULL_INTERVAL_MS)

            db.set_meta("last_sync", time.strftime("%Y-%m-%d %H:%M:%S"))
            stats["fetched"] = received
            result.update(ok=True, **stats)
            log(f"🎉 同步完成：本次新抓 {received} 篇。")
            try:
                context.close()
            except Exception:
                pass
        return result
    finally:
        db.close()


# ---------------------------------------------------------------------------
# 导出 Word
# ---------------------------------------------------------------------------
def _locate_docx(before: set):
    """在 HERE 目录里找到本次新生成（或最新）的 docx。"""
    after = set(HERE.glob("小米笔记导出_*.docx"))
    new_files = sorted(after - before, key=lambda x: x.stat().st_mtime)
    if new_files:
        return new_files[-1]
    newest = sorted(after, key=lambda x: x.stat().st_mtime)
    if newest:
        return newest[-1]
    docx_all = sorted(HERE.glob("*.docx"), key=lambda x: x.stat().st_mtime)
    return docx_all[-1] if docx_all else None


def _run_exporter_inprocess():
    """exe 环境下没有独立 python.exe，在本进程内执行外部 export_mi_notes.py。"""
    import runpy
    old_argv, old_cwd = sys.argv[:], os.getcwd()
    os.chdir(str(HERE))
    try:
        sys.argv = [str(WORD_EXPORTER), "--from-json", str(EXPORT_JSON)]
        runpy.run_path(str(WORD_EXPORTER), run_name="__main__")
    finally:
        sys.argv = old_argv
        os.chdir(old_cwd)


def export_word(log=_noop, progress=_noop, db_path: Path = DB_PATH):
    """从 db 导出 JSON，再生成 Word。返回 docx 路径。"""
    db = NotesDB(db_path)
    try:
        doc = db.export_json()
    finally:
        db.close()
    if not doc["notes"]:
        raise RuntimeError("本地还没有笔记，请先执行一次同步。")
    with open(EXPORT_JSON, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False)
    log(f"已汇总 {len(doc['notes'])} 篇笔记，开始生成 Word ...")
    progress("word", 1, 2, {})

    if WORD_EXPORTER.exists():
        before = set(HERE.glob("小米笔记导出_*.docx"))
        failed_output = None
        try:
            if FROZEN:
                # exe 内：进程内运行外部美化脚本；失败则回退内置渲染
                log("检测到 export_mi_notes.py，在程序内沿用原排版 ...")
                _run_exporter_inprocess()
            else:
                # 脚本运行：子进程调用外部美化脚本
                proc = subprocess.run(
                    [sys.executable, str(WORD_EXPORTER), "--from-json", str(EXPORT_JSON)],
                    cwd=str(HERE), capture_output=True, text=True,
                    encoding="utf-8", errors="replace")
                out = (proc.stdout or "") + (proc.stderr or "")
                m = re.search(r"([^\s\n]*小米笔记导出_[^\s\n]*\.docx)", out)
                if m and Path(m.group(1)).exists():
                    return Path(m.group(1))
                if proc.returncode != 0:
                    failed_output = out
            found = _locate_docx(before)
            if found:
                return found
            if failed_output:
                raise RuntimeError("Word 生成脚本未产出文件。\n" + failed_output[-800:])
        except RuntimeError:
            raise
        except Exception as e:
            log("外部排版脚本运行失败，改用内置排版：%s" % e)

    log("使用内置排版生成 Word ...")
    return builtin_export(doc)


def _decode_data_url(data_url: str):
    m = re.match(r"data:([^;,]*)?(;base64)?,(.*)", data_url or "", re.S)
    if not m:
        return None, None
    mime = m.group(1) or "image/jpeg"
    b64 = m.group(3)
    try:
        raw = base64.b64decode(b64)
    except Exception:
        return None, None
    return mime, raw


def _html_to_blocks(html: str):
    """极简 HTML -> [(kind, text/rawid)]。图片块以 ('img', rawId) 表示。"""
    blocks = []
    if not html:
        return blocks
    s = html
    # 图片/音视频标签 -> 占位
    media_ids = []

    def _media(mo):
        raw = mo.group(1)
        media_ids.append(raw)
        return "\n[[MEDIA:%s]]\n" % raw

    s = re.sub(r'<(?:img|sound|video)\b[^>]*?(?:id|data)=["\']([^"\']*)["\'][^>]*?\/?>',
               _media, s, flags=re.I)
    # 换行/块级闭合
    s = re.sub(r'(?i)<br\s*/?>', '\n', s)
    s = re.sub(r'(?i)</(p|div|li|h[1-6]|tr)>', '\n', s)
    s = re.sub(r'(?i)<li[^>]*>', '· ', s)
    # 去其余标签
    s = re.sub(r'<[^>]+>', '', s)
    # HTML 实体
    for a, b in (("&nbsp;", " "), ("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
                 ("&quot;", '"'), ("&#39;", "'")):
        s = s.replace(a, b)
    for line in s.split("\n"):
        line = line.strip()
        mm = re.match(r"\[\[MEDIA:(.+)\]\]", line)
        if mm:
            blocks.append(("img", mm.group(1)))
        elif line:
            blocks.append(("text", line))
    return blocks


def builtin_export(doc: dict, out_path: Path = None):
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    if out_path is None:
        out_path = HERE / ("小米笔记导出_%s.docx"
                           % time.strftime("%Y%m%d_%H%M%S"))
    folders = doc.get("folders", {})
    notes = doc.get("notes", [])
    media = doc.get("media", {})

    groups = {}
    for n in notes:
        groups.setdefault(n.get("folderId", ""), []).append(n)

    word = Document()
    style = word.styles["Normal"]
    style.font.name = "微软雅黑"
    style.font.size = Pt(11)

    # 封面
    t = word.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = t.add_run("小米笔记导出")
    run.font.size = Pt(26)
    run.bold = True
    sub = word.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run(time.strftime("%Y-%m-%d %H:%M:%S")).font.size = Pt(12)
    info = word.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info.add_run(f"共 {len(notes)} 篇 · {len(folders)} 个分组 · {len(media)} 个媒体"
                 ).font.size = Pt(10)
    word.add_page_break()

    # 未分组放最后
    def _gk(k):
        return (k == "", folders.get(k, ""))
    for fid in sorted(groups.keys(), key=lambda k: (k == "", folders.get(k, ""))):
        h = word.add_heading(folders.get(fid, "未分组") or "未分组", level=1)
        for n in sorted(groups[fid], key=lambda x: x.get("createDate", 0)):
            title = n.get("title") or time.strftime(
                "%Y-%m-%d %H:%M",
                time.localtime((n.get("createDate") or 0) / 1000))
            p = word.add_paragraph()
            tr = p.add_run(title)
            tr.bold = True
            tr.font.size = Pt(14)
            for kind, val in _html_to_blocks(n.get("content", "")):
                if kind == "text":
                    word.add_paragraph(val)
                else:
                    md = media.get(val)
                    if md and md.get("data"):
                        _mime, raw = _decode_data_url(md["data"])
                        if raw:
                            try:
                                word.add_picture(io.BytesIO(raw), width=Inches(4.5))
                            except Exception:
                                pass
            word.add_paragraph()
    word.save(str(out_path))
    return out_path


def status():
    db = NotesDB()
    try:
        return db.counts()
    finally:
        db.close()
