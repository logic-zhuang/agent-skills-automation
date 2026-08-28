# -*- coding: utf-8 -*-
"""Parse user profile text (or JSON) → internal profile.json for plan_fill.

Usage:
  python excel-com/scripts/parse_profile.py profile.txt --out profile.json
  python excel-com/scripts/parse_profile.py already.json --out profile.json --print
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

from sheet_utils import PROFILE_KEY_MAP, normalize_label

KV_LINE = re.compile(
    r"^[\s\u3000]*([^\s：:=]+)[\s\u3000]*[：:=][\s\u3000]*(.+?)\s*$"
)


def parse_kv_lines(text: str) -> dict[str, Any]:
    student: dict[str, Any] = {}
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        m = KV_LINE.match(line)
        if not m:
            continue
        label_norm = normalize_label(m.group(1))
        raw = m.group(2).strip()
        key = PROFILE_KEY_MAP.get(label_norm) or label_norm
        student[key] = raw
    return student


def parse_profile(path: Path) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8")
    stripped = text.strip()
    if path.suffix.lower() == ".json" or (stripped.startswith("{") and stripped.endswith("}")):
        data = json.loads(stripped)
        if "student" not in data and isinstance(data, dict):
            return {"student": data, "missing": []}
        data.setdefault("missing", [])
        return data

    student = parse_kv_lines(text)
    expected = ["name", "student_id"]
    missing = [k for k in expected if k not in student or not student[k]]
    return {"student": student, "missing": missing}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Parse profile text → profile.json")
    parser.add_argument("path")
    parser.add_argument("--out", metavar="profile.json")
    parser.add_argument("--print", action="store_true")
    args = parser.parse_args(argv)

    profile = parse_profile(Path(args.path))
    if args.out:
        Path(args.out).write_text(json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8")
        print("wrote", args.out)
    if args.print or not args.out:
        print(json.dumps(profile, ensure_ascii=False, indent=2))
    if profile.get("missing"):
        print("missing:", ", ".join(profile["missing"]), file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
