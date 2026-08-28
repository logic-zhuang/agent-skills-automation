# -*- coding: utf-8 -*-
"""Demo: profile text → plan → draft with TrackRevisions (review-first).

Case folder: word-com/examples/form-fill/scholarship-cnpc/
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

CASE = Path(__file__).resolve().parent
ROOT = CASE.parents[2]  # word-com/
SHARED = CASE.parent / "_shared"
sys.path.insert(0, str(ROOT / "scripts"))

SRC = Path(
    r"D:/Downloads/各类奖学金附件/各类奖学金附件/中国石油奖学金/"
    r"附件1：【文理学院-XXX】中国石油奖学金”申请审批表（电子+纸质，一式两份，双面打印）.doc"
)
FORM = CASE / "demo_scholarship_form.doc"
DRAFT = CASE / "demo_scholarship_draft.doc"
PROFILE_TXT = SHARED / "profile.txt"
PROFILE_JSON = SHARED / "profile.json"
PLAN = CASE / "demo_scholarship_plan.json"
SLOTS = CASE / "demo_scholarship_slots.json"


def main() -> int:
    if not SRC.exists() and not FORM.exists():
        print("ERROR: 找不到奖学金表。请将原件放到本目录 demo_scholarship_form.doc")
        return 1
    if SRC.exists():
        shutil.copy2(SRC, FORM)

    subprocess.check_call(
        [
            sys.executable,
            str(ROOT / "scripts" / "parse_profile.py"),
            str(PROFILE_TXT),
            "--out",
            str(PROFILE_JSON),
        ]
    )

    if DRAFT.exists():
        try:
            DRAFT.unlink()
        except PermissionError:
            alt = CASE / "demo_scholarship_draft_new.doc"
            shutil.copy2(FORM, alt)
            draft_path = alt
        else:
            shutil.copy2(FORM, DRAFT)
            draft_path = DRAFT
    else:
        shutil.copy2(FORM, DRAFT)
        draft_path = DRAFT

    from form_utils import inspect_document
    from word_session import attach_word, open_doc
    from auto_fill_form import main as auto_main

    word = attach_word(visible=True)
    doc = open_doc(word, draft_path)
    payload = inspect_document(doc)
    SLOTS.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    # draft_path is already a copy — never open/write the Downloads source
    sys.argv = [
        "auto_fill_form.py",
        str(draft_path),
        "--kb",
        str(PROFILE_JSON),
        "--slots",
        str(SLOTS),
        "--plan",
        str(PLAN),
        "--auto-suggest",
        "--save",
    ]
    auto_main()

    plan = json.loads(PLAN.read_text(encoding="utf-8"))
    print("=== 审阅演示完成（未定稿）===")
    print("case:", CASE)
    print("profile:", PROFILE_JSON)
    print("slots:", SLOTS)
    print("plan:", PLAN)
    print("draft (打开核对):", draft_path)
    print(
        f"matched={plan['matched_count']} unmatched={plan['unmatched_slot_count']} "
        f"needs_confirm={plan.get('needs_confirm_count', 0)}"
    )
    for m in plan.get("matched") or []:
        print(f"  [{m.get('fill_mode') or m.get('kind')}] {m['label']} = {m['value']}")
    print("\n用户确认后再运行:")
    print(
        f'  python word-com/scripts/auto_fill_form.py "{draft_path}" '
        f'--kb "{PROFILE_JSON}" --save'
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
