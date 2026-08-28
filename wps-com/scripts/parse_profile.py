# -*- coding: utf-8 -*-
"""Parse user profile text (or JSON) → internal profile.json for plan_fill.

Users never write slots JSON. This turns:
  姓名：李俊杰
  学号：2024216758
or a short self-intro paragraph
into {"student": {"name": "...", ...}}.

Usage:
  python wps-com/scripts/parse_profile.py demo_profile.txt --out profile.json
  python word-com/scripts/parse_profile.py already.json --out profile.json
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

from form_utils import PROFILE_KEY_MAP, normalize_label

KV_LINE = re.compile(
    r"^[\s\u3000]*([^\s：:=]+)[\s\u3000]*[：:=][\s\u3000]*(.+?)\s*$"
)


def _coerce_value(key: str, raw: str) -> Any:
    t = raw.strip()
    if key == "comprehensive_ranking":
        if t in {"是", "true", "True", "yes", "1"}:
            return True
        if t in {"否", "false", "False", "no", "0"}:
            return False
    if key == "score_rank" and "/" in t:
        return t
    return t


def parse_kv_lines(text: str) -> dict[str, Any]:
    student: dict[str, Any] = {}
    missing_hints: list[str] = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        m = KV_LINE.match(line)
        if not m:
            continue
        label, value = m.group(1), m.group(2)
        norm = normalize_label(label)
        field = PROFILE_KEY_MAP.get(norm) or PROFILE_KEY_MAP.get(label)
        if not field:
            # try partial
            for k, v in PROFILE_KEY_MAP.items():
                if k in norm or norm in k:
                    field = v
                    break
        if not field:
            missing_hints.append(label)
            continue
        student[field] = _coerce_value(field, value)
        if field == "score_rank" and "/" in value:
            a, b = value.split("/", 1)
            student.setdefault("rank", a.strip())
            student.setdefault("rank_total", b.strip())
    return student


def parse_loose_paragraph(text: str) -> dict[str, Any]:
    """Best-effort extract from free text (规则优先，不全则 missing)."""
    student: dict[str, Any] = {}
    patterns = [
        (r"(?:我叫|姓名[是为]?|名为)\s*([^\s，,。；;]{2,4})", "name"),
        (r"学号[：:\s]*([A-Za-z0-9]+)", "student_id"),
        (r"(?:学院|院系)[：:\s]*([^\s，,。；;]{2,20})", "college"),
        (r"专业[：:\s]*([^\s，,。；;]{2,20})", "major"),
        (r"(?:手机|电话|联系电话)[：:\s]*(\d{7,11})", "phone"),
        (r"身份证[号]?[：:\s]*(\d{15,18}[\dXx]?)", "id_number"),
        (r"性别[：:\s]*(男|女)", "gender"),
    ]
    for rx, field in patterns:
        m = re.search(rx, text)
        if m and field not in student:
            student[field] = m.group(1)
    return student


def parse_profile(raw: str | dict) -> dict[str, Any]:
    if isinstance(raw, dict):
        if "student" in raw:
            return {"profile": raw, "missing": [], "source": "json"}
        return {"profile": {"student": raw}, "missing": [], "source": "json_flat"}

    text = raw.strip()
    if text.startswith("{"):
        data = json.loads(text)
        return parse_profile(data)

    student = parse_kv_lines(text)
    source = "kv"
    if len(student) < 2:
        loose = parse_loose_paragraph(text)
        student = {**loose, **student}
        source = "mixed" if student else "empty"

    # Suggested fields for Agent to ask
    recommended = ["name", "student_id", "college", "phone", "id_number", "major"]
    missing = [f for f in recommended if f not in student]

    return {
        "schema": "wps-com/profile/v1",
        "profile": {"student": student},
        "missing": missing,
        "source": source,
        "message": "内部 profile；用户无需手写 JSON。缺字段由 Agent 追问。" if missing else "档案字段较完整。",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Parse profile text → profile.json")
    parser.add_argument("path", help="UTF-8 text or JSON file")
    parser.add_argument("--out", metavar="profile.json")
    parser.add_argument("--print", action="store_true", dest="print_all")
    args = parser.parse_args(argv)

    path = Path(args.path)
    raw_text = path.read_text(encoding="utf-8")
    if path.suffix.lower() == ".json":
        result = parse_profile(json.loads(raw_text))
    else:
        result = parse_profile(raw_text)

    if args.out:
        Path(args.out).write_text(
            json.dumps(result["profile"], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print("wrote", args.out)

    if args.print_all or not args.out:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(
            json.dumps(
                {"fields": len(result["profile"].get("student") or {}), "missing": result["missing"]},
                ensure_ascii=False,
            )
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
