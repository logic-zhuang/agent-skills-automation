# -*- coding: utf-8 -*-
"""Plan fill.json from inspect suggestions + profile/KB.

unknown / low confidence / table_rows layout → needs_confirm (Agent decides).

Usage:
  python excel-com/scripts/plan_fill.py slots.json --kb profile.json --out fill.json
  python excel-com/scripts/plan_fill.py slots.json --kb profile.json --print
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from sheet_utils import flatten_kb, normalize_label, resolve_kb_value


def plan_fill(inspect_payload: dict[str, Any], kb: dict[str, Any]) -> dict[str, Any]:
    slots = inspect_payload.get("suggestions") or inspect_payload.get("slots") or []
    layout_hint = (inspect_payload.get("facts") or {}).get("layout_hint") or "unknown"
    kb_flat = flatten_kb(kb)
    cells: list[dict[str, Any]] = []
    matched: list[dict[str, Any]] = []
    unmatched_slots: list[dict[str, Any]] = []
    needs_confirm: list[dict[str, Any]] = []
    used_kb_keys: set[str] = set()

    if layout_hint == "table_rows":
        needs_confirm.append(
            {
                "layout_hint": "table_rows",
                "hint": "行式表：Agent 应使用 append_row（或自写 fill ops），勿用左→右 suggestions 自动填",
            }
        )

    for slot in slots:
        label_norm = slot.get("label_norm") or normalize_label(slot.get("label", ""))
        fill_mode = slot.get("fill_mode") or "single"
        conf = float(slot.get("confidence") or 0.5)

        if layout_hint == "table_rows" or fill_mode == "unknown" or conf < 0.5:
            needs_confirm.append(
                {
                    "slot_id": slot.get("slot_id"),
                    "label": slot.get("label"),
                    "label_norm": label_norm,
                    "fill_mode": fill_mode,
                    "confidence": conf,
                    "reason": slot.get("reason"),
                    "targets": slot.get("targets"),
                    "hint": "请确认是否填写及写入地址（或改用 append_row）",
                }
            )
            continue

        kb_key, raw = resolve_kb_value(slot.get("label", ""), kb_flat)
        if kb_key is None or raw is None or raw == "":
            unmatched_slots.append(
                {
                    "slot_id": slot.get("slot_id"),
                    "label": slot.get("label"),
                    "label_norm": label_norm,
                    "fill_mode": fill_mode,
                    "sheet": slot.get("sheet"),
                    "address": slot.get("address"),
                }
            )
            continue

        text = str(raw)
        item = {
            "sheet": slot.get("sheet"),
            "address": slot.get("address"),
            "text": text,
            "slot_id": slot.get("slot_id"),
            "label": slot.get("label"),
            "op": "write_cell",
        }
        cells.append(item)
        matched.append(
            {
                "slot_id": slot.get("slot_id"),
                "label": slot.get("label"),
                "fill_mode": fill_mode,
                "confidence": conf,
                "kb_key": kb_key,
                "value": text,
                "sheet": slot.get("sheet"),
                "address": slot.get("address"),
            }
        )
        used_kb_keys.add(kb_key)

    unmatched_kb = []
    for k, v in kb_flat.items():
        if k not in used_kb_keys:
            unmatched_kb.append({"key": k, "value": v})

    fill = {"schema": "excel-com/fill/v1.1", "cells": cells, "ops": []}
    return {
        "schema": "excel-com/plan/v1.1",
        "layout_hint": layout_hint,
        "matched": matched,
        "unmatched_slots": unmatched_slots,
        "needs_confirm": needs_confirm,
        "unmatched_kb": unmatched_kb,
        "review_required": True,
        "form_fingerprint": inspect_payload.get("form_fingerprint"),
        "fill": fill,
        "agent_note": "Execute via --fill Agent plan, or --auto-suggest for high-confidence only.",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Plan Excel fill from slots + KB")
    parser.add_argument("slots_json")
    parser.add_argument("--kb", required=True)
    parser.add_argument("--out", metavar="fill_or_plan.json")
    parser.add_argument("--print", action="store_true")
    args = parser.parse_args(argv)

    inspect_payload = json.loads(Path(args.slots_json).read_text(encoding="utf-8"))
    kb = json.loads(Path(args.kb).read_text(encoding="utf-8"))
    plan = plan_fill(inspect_payload, kb)

    if args.out:
        Path(args.out).write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
        print("wrote", args.out)

    if args.print or not args.out:
        print(json.dumps(plan, ensure_ascii=False, indent=2))

    print(
        f"matched={len(plan['matched'])} unmatched_slots={len(plan['unmatched_slots'])} "
        f"needs_confirm={len(plan['needs_confirm'])} layout={plan.get('layout_hint')}",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
