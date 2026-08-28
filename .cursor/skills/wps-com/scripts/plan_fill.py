# -*- coding: utf-8 -*-
"""Plan fill.json from inspect slots + profile/KB. unknown → needs_confirm, not written.

Usage:
  python wps-com/scripts/plan_fill.py slots.json --kb kb.json --out fill.json
  python word-com/scripts/plan_fill.py slots.json --kb kb.json --print
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from form_utils import flatten_kb, normalize_label, render_slot_fill_items


def plan_fill(inspect_payload: dict[str, Any], kb: dict[str, Any]) -> dict[str, Any]:
    slots = inspect_payload.get("suggestions") or inspect_payload.get("slots") or []
    coverage = inspect_payload.get("coverage") or {}
    layout_hint = (inspect_payload.get("facts") or {}).get("layout_hint") or "unknown"
    kb_flat = flatten_kb(kb)
    cells: list[dict[str, Any]] = []
    matched: list[dict[str, Any]] = []
    unmatched_slots: list[dict[str, Any]] = []
    needs_confirm: list[dict[str, Any]] = []
    used_kb_keys: set[str] = set()

    for slot in slots:
        label_norm = slot.get("label_norm") or normalize_label(slot.get("label", ""))
        fill_mode = slot.get("fill_mode") or slot.get("kind") or "field"
        conf = float(slot.get("confidence") or 0.5)

        if fill_mode == "unknown" or conf < 0.5:
            needs_confirm.append(
                {
                    "slot_id": slot.get("slot_id"),
                    "label": slot.get("label"),
                    "label_norm": label_norm,
                    "fill_mode": fill_mode,
                    "confidence": conf,
                    "reason": slot.get("reason"),
                    "targets": slot.get("targets"),
                    "hint": "请确认整格填还是拆格填，或手填",
                }
            )
            continue

        kb_key, items, raw = render_slot_fill_items(slot, kb_flat)
        if kb_key is None or not items:
            unmatched_slots.append(
                {
                    "slot_id": slot.get("slot_id"),
                    "label": slot.get("label"),
                    "label_norm": label_norm,
                    "kind": slot.get("kind", "field"),
                    "fill_mode": fill_mode,
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
                "fill_mode": fill_mode,
                "confidence": conf,
                "kb_key": kb_key,
                "value": raw if fill_mode not in {"single", "field"} else items[0]["text"],
                "cells_written": len(items),
            }
        )
        used_kb_keys.add(kb_key)

    unmatched_kb = []
    matched_labels = {normalize_label(m["label"]) for m in matched}
    for k, v in kb_flat.items():
        if "." in k:
            continue
        if normalize_label(k) in matched_labels:
            continue
        if k in used_kb_keys or normalize_label(k) in {normalize_label(x) for x in used_kb_keys}:
            continue
        if k.startswith("_"):
            continue
        unmatched_kb.append({"key": k, "value": v})

    fill_payload: dict[str, Any] = {"cells": cells}
    if len({c["table"] for c in cells}) == 1 and cells:
        fill_payload["table"] = cells[0]["table"]

    return {
        "schema": "wps-com/fill-plan/v2.2",
        "form_fingerprint": inspect_payload.get("form_fingerprint"),
        "layout_hint": layout_hint,
        "coverage": coverage,
        "fill": fill_payload,
        "matched": matched,
        "matched_count": len(matched),
        "unmatched_slots": unmatched_slots,
        "unmatched_slot_count": len(unmatched_slots),
        "needs_confirm": needs_confirm,
        "needs_confirm_count": len(needs_confirm),
        "unmatched_kb": unmatched_kb,
        "unmatched_kb_count": len(unmatched_kb),
        "unmatched_kb_note": "档案多出来的字段；不等于漏扫。漏扫看 coverage.unresolved_labels / unbound_empties。",
        "review_required": True,
        "message": "请审阅匹配结果；用 --fill 执行 Agent 计划，或 --auto-suggest 采纳高置信建议。禁止自动 AcceptAll。",
        "agent_note": "suggestions→plan is optional; Agent may author fill.json from facts.",
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

    cov = plan.get("coverage") or {}
    summary = {
        "matched": plan["matched_count"],
        "unmatched_slots": plan["unmatched_slot_count"],
        "needs_confirm": plan["needs_confirm_count"],
        "unmatched_kb": plan["unmatched_kb_count"],
        "unresolved_labels": cov.get("unresolved_label_count"),
        "unbound_empties": cov.get("unbound_empty_count"),
        "label_coverage": cov.get("label_coverage"),
        "form_fingerprint": plan.get("form_fingerprint"),
        "review_required": True,
    }
    if args.print_plan or not args.out:
        print(json.dumps(plan, ensure_ascii=False, indent=2))
    else:
        print(json.dumps(summary, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
