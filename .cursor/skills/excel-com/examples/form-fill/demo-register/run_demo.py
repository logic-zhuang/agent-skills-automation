# -*- coding: utf-8 -*-
"""End-to-end demo for excel-com form-fill skeleton."""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL = HERE.parents[2]  # excel-com/
SCRIPTS = SKILL / "scripts"
OUT = HERE / "_out"


def run(args: list[str]) -> None:
    print("+", " ".join(args))
    subprocess.check_call(args, cwd=str(SKILL))


def main() -> int:
    OUT.mkdir(exist_ok=True)
    form = OUT / "form.xlsx"
    profile_json = OUT / "profile.json"
    slots = OUT / "slots.json"
    plan = OUT / "plan.json"
    draft = OUT / "draft.xlsx"

    py = sys.executable
    run([py, str(SCRIPTS / "create_form_template.py"), "--out", str(form)])
    run(
        [
            py,
            str(SCRIPTS / "parse_profile.py"),
            str(HERE / "profile.txt"),
            "--out",
            str(profile_json),
        ]
    )
    run([py, str(SCRIPTS / "inspect_sheet.py"), str(form), "--json", str(slots)])
    run(
        [
            py,
            str(SCRIPTS / "plan_fill.py"),
            str(slots),
            "--kb",
            str(profile_json),
            "--out",
            str(plan),
            "--print",
        ]
    )
    run(
        [
            py,
            str(SCRIPTS / "auto_fill_sheet.py"),
            str(form),
            "--kb",
            str(profile_json),
            "--plan",
            str(plan),
            "--auto-suggest",
            "--draft",
            str(draft),
            "--highlight",
        ]
    )
    print("draft:", draft)
    if plan.exists():
        data = json.loads(plan.read_text(encoding="utf-8"))
        print(
            f"matched={len(data.get('matched', []))} "
            f"needs_confirm={len(data.get('needs_confirm', []))}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
