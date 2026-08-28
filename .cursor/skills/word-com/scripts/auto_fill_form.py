# -*- coding: utf-8 -*-
"""Word form fill: Agent --fill plan or --auto-suggest; never write user source in place.

Usage (Agent-authored fill — preferred):
  python <resources_path>/scripts/auto_fill_form.py source.docx --fill fill.json --draft draft.docx

Usage (legacy heuristic suggestions):
  python <resources_path>/scripts/auto_fill_form.py source.docx --kb kb.json --auto-suggest --draft draft.docx

Plan-only (no write):
  python <resources_path>/scripts/auto_fill_form.py source.docx --kb kb.json --plan plan.json
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from fill_form import fill_form
from form_utils import inspect_document, verify_fill
from plan_fill import plan_fill
from word_session import attach_word, clone_for_edit, open_doc


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Fill Word working copy from --fill or --auto-suggest"
    )
    parser.add_argument("path", help="Source doc, or working copy when using --save only")
    parser.add_argument("--kb", help="KB / student profile JSON (required for --auto-suggest / plan)")
    parser.add_argument("--slots", metavar="slots.json", help="Reuse inspect output")
    parser.add_argument("--plan", metavar="plan.json", help="Write plan / suggestions report")
    parser.add_argument(
        "--fill",
        metavar="fill.json",
        help="Agent-authored fill plan to execute",
    )
    parser.add_argument(
        "--auto-suggest",
        action="store_true",
        help="Execute heuristic plan_fill (legacy demos)",
    )
    parser.add_argument("--hidden", action="store_true")
    parser.add_argument("--save", action="store_true", help="Save working copy")
    parser.add_argument(
        "--draft",
        metavar="OUT.docx",
        help="Clone source→OUT, open OUT only, fill",
    )
    parser.add_argument("--no-track", action="store_true")
    parser.add_argument("--no-verify", action="store_true")
    parser.add_argument("--no-save", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--inspect-source", action="store_true")
    args = parser.parse_args(argv)

    src = Path(args.path)
    do_write = bool(args.fill or args.auto_suggest)

    if do_write and not args.draft and not args.save:
        print(
            "ERROR: refuse to write source in place. Pass --draft OUT.docx "
            "(session workspace) or --save on an existing working copy.",
            file=sys.stderr,
        )
        return 2

    if args.auto_suggest and not args.kb:
        print("ERROR: --auto-suggest requires --kb", file=sys.stderr)
        return 2

    word = attach_word(visible=not args.hidden)

    inspect_payload = None
    if args.slots:
        inspect_payload = json.loads(Path(args.slots).read_text(encoding="utf-8"))
    elif (args.kb or args.plan or args.auto_suggest) and (args.inspect_source or not do_write):
        doc_src = open_doc(word, src, read_only=True)
        inspect_payload = inspect_document(doc_src)
        try:
            doc_src.Close(SaveChanges=False)
        except Exception:
            pass

    plan = None
    fill_payload = None
    if args.fill:
        raw = json.loads(Path(args.fill).read_text(encoding="utf-8"))
        fill_payload = raw.get("fill") if isinstance(raw.get("fill"), dict) else raw

    if do_write:
        if args.draft:
            work_path = Path(args.draft).resolve()
            clone_for_edit(word, src, work_path)
            doc = open_doc(word, work_path)
        else:
            work_path = src.resolve()
            doc = open_doc(word, work_path)

        if inspect_payload is None and (args.auto_suggest or args.plan):
            inspect_payload = inspect_document(doc)

        if args.auto_suggest:
            kb = json.loads(Path(args.kb).read_text(encoding="utf-8"))
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

        if not args.no_track:
            try:
                doc.TrackRevisions = True
            except Exception:
                pass

        filled = fill_form(doc, fill_payload)

        verification = None
        if not args.no_verify and fill_payload.get("cells"):
            verification = verify_fill(doc, fill_payload["cells"])
            if plan is not None:
                plan["verification"] = verification
                if args.plan:
                    Path(args.plan).write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

        saved = False
        if args.draft or (args.save and not args.no_save):
            doc.Save()
            saved = True

        result = {
            "source": str(src.resolve()),
            "working_copy": str(work_path),
            "filled": filled,
            "mode": "fill" if args.fill else "auto-suggest",
            "matched": plan["matched_count"] if plan else None,
            "unmatched_slots": plan["unmatched_slot_count"] if plan else None,
            "needs_confirm": plan.get("needs_confirm_count", 0) if plan else None,
            "layout_hint": (plan or {}).get("layout_hint")
            or (inspect_payload or {}).get("facts", {}).get("layout_hint"),
            "saved": saved,
            "track_revisions": not args.no_track,
            "review_required": True,
            "draft": str(work_path) if args.draft else None,
        }
        if verification is not None:
            result["verification"] = verification
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    # Plan-only
    if inspect_payload is None:
        doc_src = open_doc(word, src, read_only=True)
        inspect_payload = inspect_document(doc_src)
        try:
            doc_src.Close(SaveChanges=False)
        except Exception:
            pass
    if not args.kb:
        print("ERROR: plan-only mode needs --kb", file=sys.stderr)
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
        "matched": plan["matched_count"],
        "needs_confirm": plan.get("needs_confirm_count", 0),
        "plan": args.plan,
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
