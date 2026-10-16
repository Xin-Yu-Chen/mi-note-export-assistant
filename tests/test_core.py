#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mi_note_core 核心逻辑测试（不依赖网络与浏览器，无需第三方库即可跑数据库部分）。
运行：py tests/test_core.py
"""
import json
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import mi_note_core as core  # noqa: E402

TMP = Path(tempfile.mkdtemp(prefix="minote_test_"))
DB = TMP / "test.db"

# 1x1 透明 PNG（base64）
PNG_B64 = ("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk"
           "+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==")


def note(i, folder="f1", title=None, mod=1000, content="内容", with_media=False):
    n = {
        "id": "n%d" % i,
        "folderId": folder,
        "title": title or ("笔记%d" % i),
        "createDate": 1000 + i,
        "modifyDate": mod,
        "content": content,
        "files": [],
    }
    if with_media:
        n["files"] = [{"rawId": "m%d" % i}]
        n["mediaItems"] = [{"rawId": "m%d" % i, "mime": "image/png",
                            "data": "data:image/png;base64," + PNG_B64}]
    return n


def entries(*ns):
    return [{"id": n["id"], "modifyDate": n["modifyDate"]} for n in ns]


def main():
    db = core.NotesDB(DB)
    db.upsert_folders({"f1": "随笔", "f2": "诗歌"})

    # 1) 新增：空库 -> 两篇都要拉
    n1, n2 = note(1), note(2, folder="f2")
    to_fetch, deleted, stats = db.plan(entries(n1, n2))
    assert stats["new"] == 2 and stats["to_fetch"] == 2, stats
    db.upsert_note(n1)
    db.upsert_note(n2)

    # 2) 未变化：再次计划 -> 0 待抓
    _, _, stats = db.plan(entries(n1, n2))
    assert stats["to_fetch"] == 0 and stats["unchanged"] == 2, stats

    # 3) 更新：n1 修改时间变新
    n1b = note(1, mod=2000, content="改了")
    to_fetch, _, stats = db.plan(entries(n1b, n2))
    assert stats["updated"] == 1 and "n1" in to_fetch, stats
    db.upsert_note(n1b)

    # 4) 删除：云端只剩 n2，n1 被删
    _, deleted, stats = db.plan(entries(n2))
    assert stats["deleted"] == 1 and deleted == ["n1"], stats
    db.mark_deleted(deleted)
    assert db.counts()["notes"] == 1, db.counts()

    # 5) 恢复：n1 又出现在云端
    to_fetch, _, stats = db.plan(entries(n1b, n2))
    assert stats["restored"] == 1 and "n1" in to_fetch, stats
    db.upsert_note(n1b)
    assert db.counts()["notes"] == 2, db.counts()

    # 6) 媒体
    n3 = note(3, with_media=True)
    db.upsert_note(n3)
    c = db.counts()
    assert c["media"] == 1 and c["folders"] == 2, c

    # 7) 分组改名
    db.upsert_folders({"f2": "现代诗"})
    assert db.export_json()["folders"]["f2"] == "现代诗"

    # 8) 导出 JSON 往返 + 历史 JSON 导入
    doc = db.export_json()
    assert any(n["id"] == "n3" for n in doc["notes"])
    legacy = TMP / "mi_notes_data.json"
    legacy.write_text(json.dumps(doc, ensure_ascii=False), encoding="utf-8")
    db2 = core.NotesDB(TMP / "fresh.db")
    n = core.import_legacy_json(db2, legacy)
    assert n == len(doc["notes"]) and db2.counts()["media"] == 1, db2.counts()

    # 9) 内置 Word 渲染（需要 python-docx；没装则跳过）
    try:
        out = core.builtin_export(doc, TMP / "out.docx")
        assert out.exists() and out.stat().st_size > 0
        print("  [√] builtin_export 生成 docx（%.1f KB）" % (out.stat().st_size / 1024))
    except ImportError:
        print("  [跳过] 未安装 python-docx，内置导出测试跳过")

    db.close()
    db2.close()
    print("核心逻辑测试全部通过：新增/更新/删除/恢复/分组/媒体/历史导入/导出")


if __name__ == "__main__":
    main()
