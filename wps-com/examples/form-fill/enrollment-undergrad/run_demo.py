# -*- coding: utf-8 -*-
"""Demo: undergraduate enrollment form fill via WPS COM (review-first)."""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

CASE = Path(__file__).resolve().parent
ROOT = CASE.parents[2]
SHARED = CASE.parent / "_shared"
sys.path.insert(0, str(ROOT / "scripts"))

SRC = Path(r"D:/Downloads/附件1-本科新生入学登记表.docx")
FORM = CASE / "form.docx"
DRAFT = CASE / "draft.docx"
PROFILE_TXT = SHARED / "profile.txt"
PROFILE_JSON = SHARED / "profile.json"
PLAN = CASE / "plan.json"
SLOTS = CASE / "slots.json"


def main() -> int:
    if SRC.exists():
        try:
            shutil.copy2(SRC, FORM)
        except PermissionError:
            if not FORM.exists():
                raise
            print("warn: form.docx locked, reuse existing copy")
    if not FORM.exists():
        print("ERROR: missing form.docx（可从 word-com 同名用例复制，或放下载原件）")
        return 1

    subprocess.check_call(
        [
            sys.executable,
            str(ROOT / "scripts" / "parse_profile.py"),
            str(PROFILE_TXT),
            "--out",
            str(PROFILE_JSON),
        ]
    )

    draft_path = DRAFT
    if draft_path.exists():
        try:
            draft_path.unlink()
        except PermissionError:
            draft_path = CASE / "draft_new.docx"

    from form_utils import inspect_document
    from wps_session import attach_wps, open_doc
    from auto_fill_form import main as auto_main

    # Inspect blank form first (read-only facts), then auto_fill clones → draft
    wps = attach_wps(visible=True)
    doc = open_doc(wps, FORM, read_only=True)
    payload = inspect_document(doc)
    try:
        doc.Close(SaveChanges=False)
    except Exception:
        pass
    SLOTS.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    sys.argv = [
        "auto_fill_form.py",
        str(FORM),
        "--kb",
        str(PROFILE_JSON),
        "--slots",
        str(SLOTS),
        "--plan",
        str(PLAN),
        "--auto-suggest",
        "--draft",
        str(draft_path),
    ]
    auto_main()

    plan = json.loads(PLAN.read_text(encoding="utf-8"))
    cov = plan.get("coverage") or payload.get("coverage") or {}
    print("=== 本科新生入学登记表 演示（WPS，未定稿）===")
    print("case:", CASE)
    print("draft:", draft_path)
    print(
        f"matched={plan['matched_count']} unmatched={plan['unmatched_slot_count']} "
        f"needs_confirm={plan.get('needs_confirm_count', 0)}"
    )
    print(
        f"coverage: label={cov.get('label_coverage')} "
        f"unresolved={cov.get('unresolved_label_count')} "
        f"unbound_empty={cov.get('unbound_empty_count')}"
    )
    for m in plan.get("matched") or []:
        print(f"  [{m.get('fill_mode')}] {m['label']} = {m['value']}")
    if plan.get("unmatched_slots"):
        print("unmatched (slot有、档案无):")
        for u in plan["unmatched_slots"]:
            print(f"  - {u.get('label')}")
    if cov.get("unresolved_labels"):
        print("unresolved labels (漏扫候选):")
        for u in cov["unresolved_labels"][:20]:
            print(f"  - {u.get('label')} ({u.get('reason')})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
