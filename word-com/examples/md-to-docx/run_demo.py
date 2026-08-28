# -*- coding: utf-8 -*-
"""Demo: sample.md → out/sample.docx."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

CASE = Path(__file__).resolve().parent
ROOT = CASE.parents[1]  # word-com/
SRC = CASE / "sample.md"
OUT = CASE / "out" / "sample.docx"


def main() -> int:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        sys.executable,
        str(ROOT / "scripts" / "md_to_docx.py"),
        str(SRC),
        "--out",
        str(OUT),
        "--print",
    ]
    return subprocess.call(cmd)


if __name__ == "__main__":
    raise SystemExit(main())
