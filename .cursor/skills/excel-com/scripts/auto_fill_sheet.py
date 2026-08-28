# -*- coding: utf-8 -*-
"""Excel form fill: Agent fill plan or --auto-suggest; never write user source in place.

Usage (Agent-authored fill — preferred):
  python <resources_path>/scripts/auto_fill_sheet.py source.xlsx --fill fill.json --draft draft.xlsx

Usage (legacy heuristic suggestions):
  python <resources_path>/scripts/auto_fill_sheet.py source.xlsx --kb profile.json --auto-suggest --draft draft.xlsx

Plan-only (no write):
  python <resources_path>/scripts/auto_fill_sheet.py source.xlsx --kb profile.json --plan plan.json
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from excel_session import attach_excel, clone_for_edit, open_workbook
from fill_sheet import fill_sheet
from plan_fill import plan_fill
from sheet_utils import inspect_workbook, verify_fill


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Fill Excel working copy from --fill or --auto-suggest"
    )
    parser.add_argument("path", help="Source workbook, or working copy when using --save only")
    parser.add_argument("--kb", help="KB / student profile JSON (required for --auto-suggest / plan)")
    parser.add_argument("--slots", metavar="slots.json", help="Reuse inspect output")
    parser.add_argument("--sheet", default=None, help="Limit inspect to one sheet")
    parser.add_argument("--plan", metavar="plan.json", help="Write plan / suggestions report")
    parser.add_argument(
        "--fill",
        metavar="fill.json",
        help="Agent-authored fill plan to execute (cells + ops)",
    )
    parser.add_argument(
        "--auto-suggest",
        action="store_true",
        help="Execute heuristic plan_fill (legacy demos). Prefer Agent --fill for new forms.",
    )
    parser.add_argument("--hidden", action="store_true")
    parser.add_argument("--save", action="store_true", help="Save working copy (path is already draft)")
    parser.add_argument("--draft", metavar="OUT.xlsx", help="Clone source→OUT, fill OUT only")
    parser.add_argument("--highlight", action="store_true")
    parser.add_argument("--no-verify", action="store_true")
    parser.add_argument("--inspect-source", action="store_true")
    args = parser.parse_args(argv)

    src = Path(args.path)
    do_write = bool(args.fill or args.auto_suggest)

    if do_write and not args.draft and not args.save:
        print(
            "ERROR: refuse to write source in place. Pass --draft OUT.xlsx "
            "(session workspace) or --save on an existing working copy.",
            file=sys.stderr,
        )
        return 2

    if args.auto_suggest and not args.kb:
        print("ERROR: --auto-suggest requires --kb", file=sys.stderr)
        return 2

    sheet: str | int | None = args.sheet
    if sheet is not None and str(sheet).isdigit():
        sheet = int(sheet)

    excel = attach_excel(visible=not args.hidden)

    inspect_payload = None
    if args.slots:
        inspect_payload = json.loads(Path(args.slots).read_text(encoding="utf-8"))
    elif args.kb or args.plan or args.auto_suggest:
        if args.inspect_source or not do_write:
            wb_src = open_workbook(excel, src, read_only=True)
            inspect_payload = inspect_workbook(wb_src, sheet=sheet)
            try:
                wb_src.Close(SaveChanges=False)
            except Exception:
                pass

    plan = None
    fill_payload = None

    if args.fill:
        raw = json.loads(Path(args.fill).read_text(encoding="utf-8"))
        fill_payload = raw.get("fill") if isinstance(raw.get("fill"), dict) else raw
    elif args.auto_suggest or args.plan or args.kb:
        if inspect_payload is None:
            # will inspect working copy after clone, or source read-only
            pass
        if args.kb:
            kb = json.loads(Path(args.kb).read_text(encoding="utf-8"))
        else:
            kb = {}

    work_path: Path | None = None
    wb = None

    if do_write:
        if args.draft:
            work_path = Path(args.draft).resolve()
            clone_for_edit(excel, src, work_path)
            wb = open_workbook(excel, work_path)
        else:
            work_path = src.resolve()
            wb = open_workbook(excel, work_path)

        if inspect_payload is None and args.auto_suggest:
            inspect_payload = inspect_workbook(wb, sheet=sheet)

        if args.auto_suggest:
            plan = plan_fill(inspect_payload, kb)
            fill_payload = plan["fill"]
        elif fill_payload is None:
            print("ERROR: no fill payload", file=sys.stderr)
            return 2

        if plan is None and args.kb and inspect_payload is not None:
            plan = plan_fill(inspect_payload, json.loads(Path(args.kb).read_text(encoding="utf-8")))

        if args.plan and plan is not None:
            Path(args.plan).write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
            print("wrote", args.plan)

        if plan and plan.get("needs_confirm") and args.auto_suggest:
            print(
                f"WARNING: {len(plan['needs_confirm'])} need confirm — auto-suggest may write partial",
                file=sys.stderr,
            )

        filled = fill_sheet(wb, fill_payload, highlight=args.highlight)
        verification = None
        if not args.no_verify and fill_payload.get("cells"):
            verification = verify_fill(wb, fill_payload["cells"])
            if plan is not None:
                plan["verification"] = verification
                if args.plan:
                    Path(args.plan).write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

        wb.Save()
        summary = {
            "source": str(src.resolve()),
            "working_copy": str(work_path),
            "filled": filled,
            "mode": "fill" if args.fill else "auto-suggest",
            "matched": len(plan["matched"]) if plan else None,
            "needs_confirm": len(plan["needs_confirm"]) if plan else None,
            "layout_hint": (plan or {}).get("layout_hint")
            or (inspect_payload or {}).get("facts", {}).get("layout_hint"),
            "review_required": True,
            "saved": True,
            "draft": str(work_path) if args.draft else None,
            "verification": verification,
        }
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return 0

    # Plan-only: no write
    if inspect_payload is None:
        wb_src = open_workbook(excel, src, read_only=True)
        inspect_payload = inspect_workbook(wb_src, sheet=sheet)
        try:
            wb_src.Close(SaveChanges=False)
        except Exception:
            pass

    if not args.kb:
        print("ERROR: plan-only mode needs --kb (or pass --fill/--auto-suggest to write)", file=sys.stderr)
        return 2
    kb = json.loads(Path(args.kb).read_text(encoding="utf-8"))
    plan = plan_fill(inspect_payload, kb)
    if args.plan:
        Path(args.plan).write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
        print("wrote", args.plan)
    print(json.dumps({
        "wrote": False,
        "reason": "pass --fill fill.json or --auto-suggest to write a draft",
        "layout_hint": plan.get("layout_hint"),
        "matched": len(plan["matched"]),
        "needs_confirm": len(plan["needs_confirm"]),
        "plan": args.plan,
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
