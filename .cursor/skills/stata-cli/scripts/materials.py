"""Look up CLI command dictionary entries without dumping the whole YAML."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

_SKILL_ROOT = Path(__file__).resolve().parents[1]
MATERIALS_PATH = _SKILL_ROOT / "materials.yaml"

_ENTRY_SPLIT = re.compile(r"(?m)^  - id: ")
_FIELD = re.compile(r"(?m)^    ([a-z_]+):\s*(.*)$")


def materials_path() -> Path:
    return MATERIALS_PATH


def _parse_entry(block: str) -> dict[str, Any]:
    rec: dict[str, Any] = {}
    lines = block.splitlines()
    if not lines:
        return rec
    rec["id"] = lines[0].strip()
    key = None
    buf: list[str] = []
    mode = ""
    for ln in lines[1:]:
        m = _FIELD.match(ln)
        if m and not (mode and ln.startswith("      ")):
            if key:
                rec[key] = _join(buf, mode)
            key = m.group(1)
            rest = m.group(2).strip()
            buf = []
            if rest == ">":
                mode = "fold"
            elif rest == "|":
                mode = "literal"
            else:
                mode = ""
                rec[key] = rest.strip().strip('"')
                key = None
            continue
        if key is not None:
            buf.append(ln)
    if key:
        rec[key] = _join(buf, mode)
    tags = rec.get("tags") or ""
    if isinstance(tags, str) and tags.startswith("[") and tags.endswith("]"):
        rec["tags"] = [t.strip() for t in tags[1:-1].split(",") if t.strip()]
    return rec


def _join(buf: list[str], mode: str) -> str:
    parts: list[str] = []
    for ln in buf:
        if ln.startswith("      "):
            parts.append(ln[6:])
        else:
            parts.append(ln.strip())
    text = "\n".join(parts).strip()
    if mode == "fold":
        return " ".join(text.split())
    return text


def load_entries(path: Path | None = None) -> list[dict[str, Any]]:
    raw = (path or MATERIALS_PATH).read_text(encoding="utf-8")
    chunks = _ENTRY_SPLIT.split(raw)
    out: list[dict[str, Any]] = []
    for chunk in chunks[1:]:
        rec = _parse_entry(chunk)
        if rec.get("id"):
            out.append(rec)
    return out


def family_index(entries: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    rows = entries if entries is not None else load_entries()
    families: dict[str, int] = {}
    for rec in rows:
        fam = str(rec.get("family") or "other")
        families[fam] = families.get(fam, 0) + 1
    return {
        "count": len(rows),
        "families": [{"id": k, "n": families[k]} for k in sorted(families)],
        "ids": [str(r["id"]) for r in rows],
    }


def search_materials(
    query: str,
    *,
    family: str = "",
    limit: int = 6,
    path: Path | None = None,
) -> dict[str, Any]:
    q = (query or "").strip().lower()
    fam = (family or "").strip().lower()
    if not q and not fam:
        return {"error": "query or family is required", "hint": "Try query=regress"}
    rows = load_entries(path)
    scored: list[tuple[int, dict[str, Any]]] = []
    for rec in rows:
        if fam and str(rec.get("family") or "").lower() != fam:
            continue
        rid = str(rec.get("id") or "").lower()
        summary = str(rec.get("summary") or "").lower()
        api = str(rec.get("api") or "").lower()
        blob = " ".join(
            str(rec.get(k) or "")
            for k in ("id", "summary", "api", "notes", "example", "tags")
        ).lower()
        if q and q not in blob:
            continue
        if not q:
            score = 2
        elif rid == q or rid == f"cmd.{q}" or rid.endswith("." + q):
            score = 0
        elif q in rid:
            score = 1
        elif q in summary or q in api:
            score = 2
        else:
            score = 3
        scored.append(
            (
                score,
                {
                    "id": rec.get("id"),
                    "family": rec.get("family"),
                    "summary": rec.get("summary"),
                    "api": rec.get("api"),
                    "notes": rec.get("notes"),
                    "example": rec.get("example"),
                    "tags": rec.get("tags") or [],
                },
            )
        )
    scored.sort(key=lambda x: x[0])
    cap = max(1, min(int(limit or 6), 20))
    hits = [row for _, row in scored[:cap]]
    return {
        "query": query,
        "family": family or None,
        "hits": hits,
        "n": len(hits),
        "hint": "Matching materials entries only — do not dump materials.yaml.",
    }
