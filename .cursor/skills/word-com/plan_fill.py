# -*- coding: utf-8 -*-
"""Plan fill.json from inspect slots + OpenWorker KB data.

Usage:
  python word-com/scripts/plan_fill.py slots.json --kb kb.json --out fill.json
  python word-com/scripts/plan_fill.py slots.json --kb kb.json --print

kb.json (nested or flat):
  {"student": {"name": "张三", "student_id": "2021001", "college": "石油学院"}}
  or {"姓名": "张三", "学号": "2021001"}
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from form_utils import flatten_kb, normalize_label, render_slot_fill_items


def plan_fill(inspect_payload: dict[str, Any], kb: dict[str, Any]) -> dict[str, Any]:
    slots = inspect_payload.get("slots") or []
    if not slots and inspect_payload.get("table"):
        slots = inspect_payload.get("slots") or []

    kb_flat = flatten_kb(kb)
    cells: list[dict[str, Any]] = []
    matched: list[dict[str, Any]] = []
    unmatched_slots: list[dict[str, Any]] = []
    used_kb_keys: set[str] = set()

    for slot in slots:
        label_norm = slot.get("label_norm") or normalize_label(slot.get("label", ""))
        kb_key, items, raw = render_slot_fill_items(slot, kb_flat)
        if kb_key is None or not items:
            unmatched_slots.append(
                {
                    "slot_id": slot.get("slot_id"),
                    "label": slot.get("label"),
                    "label_norm": label_norm,
                    "kind": slot.get("kind", "field"),
                    "table": slot.get("table"),
                    "row": slot.get("row"),
                    "col": slot.get("col"),
                }
            )
            continue

        cells.extend(items)
        matched.append(
            {
                "slot_id": slot.get("slot_id"),
                "label": slot.get("label"),
                "kind": slot.get("kind", "field"),
                "kb_key": kb_key,
                "value": raw if slot.get("kind") != "field" else items[0]["text"],
                "cells_written": len(items),
            }
        )
        used_kb_keys.add(kb_key)

    unmatched_kb = []
    for k, v in kb_flat.items():
        if "." in k:
            continue
        if normalize_label(k) in {normalize_label(m["label"]) for m in matched}:
            continue
        if k in used_kb_keys or normalize_label(k) in {normalize_label(x) for x in used_kb_keys}:
            continue
        if k.startswith("_"):
            continue
        unmatched_kb.append({"key": k, "value": v})

    fill_payload = {"cells": cells}
    if len({c["table"] for c in cells}) == 1 and cells:
        fill_payload["table"] = cells[0]["table"]

    return {
        "schema": "word-com/fill-plan/v1",
        "form_fingerprint": inspect_payload.get("form_fingerprint"),
        "fill": fill_payload,
        "matched": matched,
        "matched_count": len(matched),
        "unmatched_slots": unmatched_slots,
        "unmatched_slot_count": len(unmatched_slots),
        "unmatched_kb": unmatched_kb,
        "unmatched_kb_count": len(unmatched_kb),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Plan fill.json from slots + KB")
    parser.add_argument("slots", help="inspect_form output JSON")
    parser.add_argument("--kb", required=True, help="OpenWorker KB / student profile JSON")
    parser.add_argument("--out", metavar="fill.json", help="Write fill.json for fill_form.py")
    parser.add_argument("--print", action="store_true", dest="print_plan", help="Print full plan report")
    args = parser.parse_args(argv)

    inspect_payload = json.loads(Path(args.slots).read_text(encoding="utf-8"))
    kb = json.loads(Path(args.kb).read_text(encoding="utf-8"))
    plan = plan_fill(inspect_payload, kb)

    if args.out:
        Path(args.out).write_text(
            json.dumps(plan["fill"], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print("wrote", args.out)

    summary = {
        "matched": plan["matched_count"],
        "unmatched_slots": plan["unmatched_slot_count"],
        "unmatched_kb": plan["unmatched_kb_count"],
        "form_fingerprint": plan.get("form_fingerprint"),
    }
    if args.print_plan or not args.out:
        print(json.dumps(plan, ensure_ascii=False, indent=2))
    else:
        print(json.dumps(summary, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
