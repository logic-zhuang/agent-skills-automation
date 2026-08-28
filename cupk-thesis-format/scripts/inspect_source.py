# -*- coding: utf-8 -*-
"""Dump a user thesis Word file for CUPK 套版: outline, cover hints, counts.

Usage:
  python <resources_path>/scripts/inspect_source.py path/to.docx
  python <resources_path>/scripts/inspect_source.py path/to.docx --json
  python <resources_path>/scripts/inspect_source.py path/to.docx --extract out.json
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import win32com.client as win32

HEADING_STYLES = {"标题 1", "标题 2", "标题 3", "Heading 1", "Heading 2", "Heading 3"}
H1_STYLES = {"标题 1", "Heading 1"}
HEADING_RE = re.compile(
    r"^(第[0-9一二三四五六七八九十]+章|摘\s*要|ABSTRACT|前\s*言|绪\s*论|结\s*论|参\s*考\s*文\s*献|致\s*谢|附录)"
)
MASTER_MARKERS = (
    "中图分类号",
    "单位代码",
    "UDC",
    "密级",
    "学号",
    "第二导师",
    "学习方式",
    "硕士学位论文",
    "硕士专业学位论文",
)


def attach_word():
    try:
        word = win32.GetActiveObject("Word.Application")
    except Exception:
        word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    return word


def open_readonly(word, path: Path):
    path = path.resolve()
    target = str(path).lower()
    for i in range(1, word.Documents.Count + 1):
        d = word.Documents(i)
        try:
            if str(Path(d.FullName).resolve()).lower() == target:
                return d, False
        except Exception:
            continue
    return word.Documents.Open(str(path), ReadOnly=True), True


def para_text(p) -> str:
    return (p.Range.Text or "").replace("\r", "").replace("\x07", "")


def style_name(p) -> str:
    try:
        return str(p.Style.NameLocal)
    except Exception:
        return "?"


def is_h1(st: str, t: str) -> bool:
    if st.startswith("TOC") or st.startswith("目录"):
        return False
    if st in H1_STYLES:
        return bool(t.strip())
    return bool(
        t
        and HEADING_RE.match(t)
        and st not in ("标题 2", "标题 3", "Heading 2", "Heading 3")
    )


def extract_cover_table(doc) -> dict:
    if doc.Tables.Count < 1:
        return {}
    t = doc.Tables(1)
    try:
        if t.Rows.Count < 5 or t.Columns.Count < 2:
            return {}
    except Exception:
        return {}

    def cell(r, c):
        try:
            return (t.Cell(r, c).Range.Text or "").replace("\r", "").replace("\x07", "").strip()
        except Exception:
            return ""

    cover = {
        "title": cell(1, 2),
        "title_cont": cell(2, 2),
        "college": cell(3, 2),
        "major": cell(4, 2),
        "name": cell(5, 2),
        "advisor": cell(6, 2) if t.Rows.Count >= 6 else "",
    }
    return {k: v for k, v in cover.items() if v}


def extract_chapters(doc):
    chapters = []
    current = None
    preamble = []
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        t = para_text(p)
        st = style_name(p)
        if is_h1(st, t):
            if current:
                chapters.append(current)
            current = {"heading": t, "style": st, "i": i, "blocks": []}
            continue
        if st.startswith("TOC") or st.startswith("目录"):
            continue
        style = "正文"
        if st in ("标题 2", "Heading 2"):
            style = "标题 2"
        elif st in ("标题 3", "Heading 3"):
            style = "标题 3"
        block = {"style": style, "text": t}
        if current is None:
            preamble.append(block)
        else:
            current["blocks"].append(block)
    if current:
        chapters.append(current)
    return preamble, chapters


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("path")
    parser.add_argument("--json", action="store_true")
    parser.add_argument(
        "--extract",
        metavar="OUT.json",
        help="write chapters/abstract dump for convert_ops fill-section",
    )
    args = parser.parse_args(argv)

    path = Path(args.path)
    if path.name.startswith("~$"):
        print("skip Word lock file", path, file=sys.stderr)
        return 2
    if not path.exists():
        print("not found", path, file=sys.stderr)
        return 2

    word = attach_word()
    owned = False
    try:
        doc, owned = open_readonly(word, path)
        info = {
            "path": str(path.resolve()),
            "paragraphs": doc.Paragraphs.Count,
            "sections": doc.Sections.Count,
            "tables": doc.Tables.Count,
            "comments": doc.Comments.Count,
            "shapes": doc.Shapes.Count,
            "inline_shapes": doc.InlineShapes.Count,
            "omaths": doc.OMaths.Count,
        }
        try:
            info["pages"] = doc.ComputeStatistics(2)
        except Exception:
            info["pages"] = None

        headings = []
        keywords = []
        master_hits = []
        cover_hints = []
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            t = para_text(p)
            st = style_name(p)
            compact = t.replace(" ", "")
            if st in HEADING_STYLES or (t and HEADING_RE.match(t)):
                headings.append({"i": i, "style": st, "text": t[:120]})
            if t.startswith("关键词") or t.lower().startswith("key word"):
                keywords.append({"i": i, "text": t[:200]})
            for m in MASTER_MARKERS:
                if m in t:
                    master_hits.append({"i": i, "marker": m, "text": t[:80]})
                    break
            if i <= 60 and t.strip() and any(
                k in compact for k in ("题目", "学院", "专业", "指导教师", "学生姓名", "起止时间")
            ):
                cover_hints.append({"i": i, "text": t[:80]})

        payload = {
            "info": info,
            "headings": headings,
            "keywords": keywords,
            "cover_hints": cover_hints,
            "master_fields_to_drop": master_hits,
        }
        if args.extract:
            preamble, chapters = extract_chapters(doc)
            extracted = {
                "path": info["path"],
                "cover_table": extract_cover_table(doc),
                "cover_hints": cover_hints,
                "keywords": keywords,
                "master_fields_to_drop": master_hits,
                "preamble": preamble,
                "chapters": chapters,
            }
            out = Path(args.extract)
            out.write_text(json.dumps(extracted, ensure_ascii=False, indent=2), encoding="utf-8")
            print("wrote extract", out, "chapters=", len(chapters))
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            print("===", path.name, "===")
            for k, v in info.items():
                print(f"  {k}: {v}")
            print("--- cover hints ---")
            for row in cover_hints:
                print(f"  {row['i']:4d} {row['text']}")
            print("--- keywords ---")
            for row in keywords:
                print(f"  {row['i']:4d} {row['text']}")
            print("--- outline ---")
            for h in headings:
                print(f"  {h['i']:4d} [{h['style']}] {h['text']}")
            if master_hits:
                print("--- drop (硕士字段) ---")
                for row in master_hits:
                    print(f"  {row['i']:4d} {row['marker']}: {row['text']}")
        if owned:
            doc.Close(False)
    finally:
        if owned:
            try:
                if word.Documents.Count == 0:
                    word.Quit()
            except Exception:
                pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
