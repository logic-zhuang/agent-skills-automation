# -*- coding: utf-8 -*-
"""Map inspect_source extract.json to fill-section JSON files."""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent
EXAMPLE = Path(__file__).resolve().parents[2] / "example" / "cupk-phase3"


def norm_heading(text: str) -> str:
    t = re.sub(r"\s+", " ", text.strip())
    if t.startswith("第") and "章" in t and "  " not in t:
        t = re.sub(r"^(第\d+章)\s*", r"\1  ", t)
    return t


def chapter_payload(ch: dict) -> dict:
    heading = norm_heading(ch["heading"])
    payload = {
        "heading": heading,
        "even_header": heading,
        "blocks": [],
    }
    for block in ch.get("blocks") or []:
        text = (block.get("text") or "").strip()
        if not text:
            continue
        if re.match(r"^k=.*\(\d+\.\d+\)$", text.replace(" ", "")):
            linear, num = text.rsplit("(", 1)
            payload["blocks"].append(
                {
                    "type": "omath",
                    "linear": linear.strip(),
                    "number": f"({num}",
                }
            )
            continue
        payload["blocks"].append({"style": block.get("style") or "正文", "text": text})
    return payload


def find_chapter(chapters: list, *patterns: str):
    for ch in chapters:
        h = ch.get("heading", "")
        if any(re.search(p, h) for p in patterns):
            return ch
    return None


def main() -> int:
    extract = json.loads((ROOT / "extract.json").read_text(encoding="utf-8"))
    chapters = extract["chapters"]
    EXAMPLE.mkdir(parents=True, exist_ok=True)

    cover = {
        "title": "Phase3 套版测试论文",
        "title_cont": "",
        "college": "石油学院",
        "major": "石油工程",
        "name": "测试同学",
        "advisor": "张老师 副教授",
        "date": "起止时间：2025 年 9 月 1 日 至 2026 年 6 月 15 日",
    }
    (EXAMPLE / "cover.json").write_text(json.dumps(cover, ensure_ascii=False, indent=2), encoding="utf-8")

    summary = find_chapter(chapters, r"摘\s*要")
    if summary:
        body = [b for b in summary["blocks"] if not (b.get("text") or "").startswith("关键词")]
        payload = {
            "blocks": [{"style": b.get("style") or "正文", "text": b.get("text", "")} for b in body],
            "keywords": "套版测试；Word COM；本科论文",
        }
        (EXAMPLE / "summary.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    abstract = find_chapter(chapters, r"ABSTRACT")
    if abstract:
        body = [b for b in abstract["blocks"] if not (b.get("text") or "").lower().startswith("key word")]
        payload = {
            "title_en": "Phase 3 Pipeline Test Thesis",
            "blocks": [{"style": b.get("style") or "正文", "text": b.get("text", "")} for b in body],
            "key_words": "formatting test; Word COM; undergraduate thesis",
        }
        (EXAMPLE / "abstract_en.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    mapping = [
        ("preface.json", r"前\s*言", "前    言", 5),
        ("ch1.json", r"第1章", "第1章  绪论", 6),
        ("ch2.json", r"第2章", "第2章  研究方法", 7),
        ("conclusion.json", r"结\s*论", "结    论", 11),
        ("refs.json", r"参\s*考\s*文\s*献", "参考文献", 12),
        ("thanks.json", r"致\s*谢", "致    谢", 14),
    ]
    manifest = []
    for fname, pattern, heading, section in mapping:
        ch = find_chapter(chapters, pattern)
        if not ch:
            continue
        payload = chapter_payload(ch)
        payload["heading"] = heading
        payload["even_header"] = heading
        (EXAMPLE / fname).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        manifest.append({"file": fname, "section": section, "heading": heading})

    (EXAMPLE / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print("wrote fill JSON to", EXAMPLE)
    for row in manifest:
        print(f"  section {row['section']:2d} <- {row['file']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
