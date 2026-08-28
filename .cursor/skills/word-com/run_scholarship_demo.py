# -*- coding: utf-8 -*-
"""One-shot demo: copy scholarship form + auto fill with demo_kb.json."""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

SRC = Path(
    r"D:/Downloads/各类奖学金附件/各类奖学金附件/中国石油奖学金/"
    r"附件1：【文理学院-XXX】中国石油奖学金”申请审批表（电子+纸质，一式两份，双面打印）.doc"
)
FORM = ROOT / "examples" / "demo_scholarship_form.doc"
OUT = ROOT / "examples" / "demo_scholarship_filled.doc"
KB = ROOT / "examples" / "demo_kb.json"
PLAN = ROOT / "examples" / "demo_scholarship_plan.json"
SLOTS = ROOT / "examples" / "demo_scholarship_slots.json"


def main() -> int:
    if not SRC.exists():
        print("ERROR: 找不到奖学金表，请把原件放到 examples/demo_scholarship_form.doc")
        return 1

    shutil.copy2(SRC, FORM)
    out_path = OUT
    if out_path.exists():
        try:
            out_path.unlink()
        except PermissionError:
            out_path = ROOT / "examples" / "demo_scholarship_filled_v2.doc"
    shutil.copy2(FORM, out_path)

    from word_session import attach_word, open_doc
    from form_utils import inspect_document
    from plan_fill import plan_fill
    from fill_form import fill_form

    word = attach_word(visible=True)
    doc = open_doc(word, out_path)

    inspect_payload = inspect_document(doc)
    SLOTS.write_text(json.dumps(inspect_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    kb = json.loads(KB.read_text(encoding="utf-8"))
    plan = plan_fill(inspect_payload, kb)
    PLAN.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    filled = fill_form(doc, plan["fill"])
    doc.Save()

    print("=== 演示完成 ===")
    print("知识库:", KB)
    print("空白表:", FORM)
    print("已填表:", out_path)
    print("slots:", SLOTS)
    print("plan:", PLAN)
    print(f"匹配: {plan['matched_count']}  未匹配: {plan['unmatched_slot_count']}  写入: {filled}")
    print("\n已填字段:")
    for m in plan["matched"]:
        kind = m.get("kind", "field")
        print(f"  [{kind}] {m['label']} = {m['value']}")
    print("\n请在 Word 窗口查看:", out_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
