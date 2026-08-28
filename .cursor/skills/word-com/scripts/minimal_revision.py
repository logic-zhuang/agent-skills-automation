# -*- coding: utf-8 -*-
"""Minimal-diff Word revisions via COM (TrackRevisions + SequenceMatcher).

LLM may return a whole paragraph; this script writes back only changed spans so
Word shows fine-grained insert/delete marks—not whole-paragraph delete+insert.

Usage:
  python word-com/scripts/minimal_revision.py path/to.docx --paragraph 2 --new "润色后的第2段"
  python word-com/scripts/minimal_revision.py path/to.docx --job revise.json
  python word-com/scripts/minimal_revision.py --active --replace "旧句" "新句"

revise.json:
  {"paragraph": 2, "old": "...", "new": "...", "expected_old": "..."}
  or {"document": true, "old": "...", "new": "...", "expected_old": "..."}
  or {"replacements": [{"old": "A", "new": "B"}, ...]}
"""
from __future__ import annotations

import argparse
import json
import sys
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import win32com.client as win32

wdCollapseEnd = 0


class MinimalRevisionError(RuntimeError):
    pass


def attach_word(*, visible: bool = True) -> Any:
    try:
        word = win32.GetActiveObject("Word.Application")
    except Exception:
        word = win32.DispatchEx("Word.Application")
    word.Visible = bool(visible)
    word.DisplayAlerts = 0
    try:
        word.ScreenUpdating = True
    except Exception:
        pass
    return word


def find_open_doc(word: Any, path: Path) -> Any | None:
    target = str(path.resolve()).lower()
    for i in range(1, word.Documents.Count + 1):
        d = word.Documents(i)
        try:
            if str(Path(d.FullName).resolve()).lower() == target:
                return d
        except Exception:
            continue
    return None


def open_doc(word: Any, path: Path) -> Any:
    path = path.resolve()
    if not path.exists():
        raise FileNotFoundError(path)
    doc = find_open_doc(word, path)
    if doc is None:
        doc = word.Documents.Open(str(path), ReadOnly=False)
    try:
        word.Activate()
        doc.Activate()
    except Exception:
        pass
    return doc


def connect_active() -> tuple[Any, Any]:
    try:
        word = win32.GetActiveObject("Word.Application")
    except Exception as exc:
        raise MinimalRevisionError("找不到正在运行的 Word，请先打开 Word。") from exc
    try:
        doc = word.ActiveDocument
        if doc is None:
            raise MinimalRevisionError("Word 当前没有活动文档。")
    except MinimalRevisionError:
        raise
    except Exception as exc:
        raise MinimalRevisionError("无法读取 Word 的活动文档。") from exc
    return word, doc


def paragraph_body_text(paragraph: Any) -> str:
    return str(paragraph.Range.Text).rstrip("\r\x07")


def get_paragraph(doc: Any, paragraph_number: int) -> Any:
    return doc.Paragraphs(paragraph_number)


def enable_tracking(doc: Any) -> None:
    doc.TrackRevisions = True


def replace_range_minimal(doc: Any, match: Any, old: str, new: str) -> None:
    """Apply old→new on match range using character-level diff (back to front)."""
    if old == new:
        return
    base = int(match.Start)
    matcher = SequenceMatcher(None, old, new, autojunk=False)
    for tag, old_start, old_end, new_start, new_end in reversed(matcher.get_opcodes()):
        if tag == "equal":
            continue
        target = doc.Range(base + old_start, base + old_end)
        replacement = new[new_start:new_end]
        if tag == "delete":
            target.Delete()
        elif tag == "insert":
            target.InsertBefore(replacement)
        else:
            target.Text = replacement


def find_text(doc: Any, text: str, *, start: int | None = None) -> Any | None:
    if not text:
        raise ValueError("text 不能为空")
    end = doc.Content.End
    search_range = doc.Range(start or doc.Content.Start, end)
    find = search_range.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    find.Text = text
    find.Forward = True
    find.Wrap = 0
    find.MatchCase = True
    find.MatchWholeWord = False
    return search_range if find.Execute() else None


def replace_text(
    doc: Any,
    old: str,
    new: str,
    *,
    all_occurrences: bool = False,
) -> int:
    if not old:
        raise ValueError("old 不能为空")
    enable_tracking(doc)
    count = 0
    search_start = doc.Content.Start
    while True:
        match = find_text(doc, old, start=search_start)
        if match is None:
            return count
        replace_range_minimal(doc, match, old, new)
        count += 1
        if not all_occurrences:
            return count
        search_start = match.End


def replace_text_in_paragraph(
    doc: Any,
    paragraph_number: int,
    old: str,
    new: str,
) -> int:
    if paragraph_number < 1:
        raise ValueError("paragraph_number 必须从 1 开始")
    if not old:
        raise ValueError("old 不能为空")
    try:
        paragraph_range = get_paragraph(doc, paragraph_number).Range.Duplicate
    except Exception as exc:
        raise IndexError(f"找不到第 {paragraph_number} 段") from exc

    find = paragraph_range.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    find.Text = old
    find.Forward = True
    find.Wrap = 0
    find.MatchCase = True
    find.MatchWholeWord = False
    if not find.Execute():
        return 0
    enable_tracking(doc)
    replace_range_minimal(doc, paragraph_range, old, new)
    return 1


def replace_document_text(
    doc: Any,
    new_text: str,
    *,
    expected_old_text: str | None = None,
) -> None:
    content = doc.Content.Duplicate
    target = doc.Range(content.Start, content.End - 1)
    old_text = str(target.Text)
    normalized_new = new_text.replace("\r\n", "\r").replace("\n", "\r")
    if expected_old_text is not None:
        normalized_expected = expected_old_text.replace("\r\n", "\r").replace("\n", "\r")
        if normalized_expected.rstrip("\r") != old_text.rstrip("\r"):
            raise MinimalRevisionError(
                "Word 当前正文与 expected_old 不一致，已停止写回（可能用户期间改过文档）。"
            )
    enable_tracking(doc)
    replace_range_minimal(doc, target, old_text, normalized_new)


def revise_paragraph_by_number(
    doc: Any,
    paragraph_number: int,
    new_text: str,
    *,
    expected_old: str | None = None,
) -> int:
    paragraph = get_paragraph(doc, paragraph_number)
    old_text = paragraph_body_text(paragraph)
    if expected_old is not None and expected_old != old_text:
        raise MinimalRevisionError(
            f"第 {paragraph_number} 段与 expected_old 不一致，已停止写回。"
        )
    return replace_text_in_paragraph(doc, paragraph_number, old_text, new_text)


def run_job(doc: Any, job: dict) -> dict:
    """Execute one revision job dict; return summary."""
    if "replacements" in job:
        total = 0
        for item in job["replacements"]:
            total += replace_text(doc, item["old"], item["new"], all_occurrences=bool(item.get("all")))
        return {"replacements": total, "revisions": doc.Revisions.Count}

    old = job.get("old", "")
    new = job.get("new", "")
    expected = job.get("expected_old")

    if job.get("document"):
        if not new and old:
            raise ValueError("document job 需要 new")
        if expected is None and old:
            expected = old
        replace_document_text(doc, new, expected_old_text=expected)
        return {"scope": "document", "revisions": doc.Revisions.Count}

    para = job.get("paragraph")
    if para is not None:
        if not old:
            old = paragraph_body_text(get_paragraph(doc, int(para)))
        if expected is not None and expected != old:
            raise MinimalRevisionError(f"第 {para} 段 expected_old 与 old 不一致")
        count = replace_text_in_paragraph(doc, int(para), old, new)
        return {"paragraph": para, "replacements": count, "revisions": doc.Revisions.Count}

    count = replace_text(doc, old, new, all_occurrences=bool(job.get("all")))
    return {"replacements": count, "revisions": doc.Revisions.Count}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Minimal-diff Word revisions (TrackRevisions)")
    parser.add_argument("path", nargs="?", help=".docx path (omit with --active)")
    parser.add_argument("--active", action="store_true", help="use Word ActiveDocument")
    parser.add_argument("--hidden", action="store_true")
    parser.add_argument(
        "--no-save",
        action="store_true",
        help="Do not Save after revision (default: Save so next agent step sees persisted edits)",
    )
    parser.add_argument(
        "--draft",
        metavar="OUT.docx",
        help="SaveAs a draft copy after revision (still leaves TrackRevisions on)",
    )
    parser.add_argument("--paragraph", type=int, help="1-based paragraph to revise")
    parser.add_argument("--new", dest="new_text", help="revised paragraph/document text")
    parser.add_argument("--old", dest="old_text", help="explicit old text (default: read from paragraph)")
    parser.add_argument("--expected-old", help="abort if current text differs")
    parser.add_argument("--replace", nargs=2, metavar=("OLD", "NEW"), action="append", help="one replacement")
    parser.add_argument("--all", action="store_true", help="replace all occurrences (with --replace)")
    parser.add_argument("--job", help="JSON job file path")
    args = parser.parse_args(argv)

    owned = False
    try:
        if args.active:
            word, doc = connect_active()
        else:
            if not args.path:
                parser.error("path required unless --active")
            word = attach_word(visible=not args.hidden)
            doc = open_doc(word, Path(args.path))
            owned = True

        if args.job:
            job = json.loads(Path(args.job).read_text(encoding="utf-8"))
            summary = run_job(doc, job)
        elif args.paragraph is not None:
            if not args.new_text:
                parser.error("--paragraph 需要 --new")
            if args.old_text:
                if args.expected_old and args.expected_old != args.old_text:
                    raise MinimalRevisionError("expected-old 与 --old 不一致")
                count = replace_text_in_paragraph(doc, args.paragraph, args.old_text, args.new_text)
            else:
                count = revise_paragraph_by_number(
                    doc,
                    args.paragraph,
                    args.new_text,
                    expected_old=args.expected_old,
                )
            summary = {"paragraph": args.paragraph, "replacements": count}
        elif args.replace:
            total = 0
            for old, new in args.replace:
                total += replace_text(doc, old, new, all_occurrences=args.all)
            summary = {"replacements": total}
        else:
            parser.error("need --job, --paragraph, or --replace")

        summary["revisions"] = doc.Revisions.Count
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        print("TrackRevisions=True; 请在 Word 审阅中接受/拒绝。未自动 AcceptAll。")
        if args.draft:
            draft = Path(args.draft).resolve()
            doc.SaveAs2(str(draft))
            print("draft", draft)
        elif not args.no_save:
            doc.Save()
            print("saved", doc.FullName)
        else:
            print("not saved (--no-save); next agent step may not see edits until you Save in Word.")
        return 0
    except (MinimalRevisionError, ValueError, IndexError, FileNotFoundError) as exc:
        print("error:", exc, file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
