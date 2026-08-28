# -*- coding: utf-8 -*-
"""Fragile CUPK 套版 operations. Never rewrite PAGE fields.

Usage:
  python <resources_path>/scripts/convert_ops.py clone --out "输出.docx"
  python <resources_path>/scripts/convert_ops.py cover --doc OUT.docx --json cover.json
  python <resources_path>/scripts/convert_ops.py fill-section --doc OUT.docx --section 6 --json ch1.json
  python <resources_path>/scripts/convert_ops.py extra-chapter --doc OUT.docx --before 11 --json ch6.json
  python <resources_path>/scripts/convert_ops.py delete-section --doc OUT.docx --index 13
  python <resources_path>/scripts/convert_ops.py finish --doc OUT.docx
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path

from cupk_com import (
    CM,
    ODD_HEADER,
    apply_heading1,
    apply_keywords_en_format,
    apply_keywords_zh_format,
    append_in_section,
    attach_word,
    center_all_heading1,
    clear_after_para,
    ensure_odd_header,
    fillable_template_path,
    open_or_reuse,
    para_text,
    replace_cell_text_keep_format,
    set_even_header_text,
    set_para_text,
    unlink_headers_keep_footers,
    wdAlignParagraphCenter,
    wdCharacter,
    wdSectionBreakNextPage,
)

def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def cmd_clone(args) -> int:
    try:
        src = Path(args.template) if args.template else fillable_template_path()
    except FileNotFoundError as e:
        print(e, file=sys.stderr)
        return 2
    dst = Path(args.out)
    if not src.exists():
        print("template not found:", src, file=sys.stderr)
        return 2
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists() and not args.force:
        print("exists (use --force):", dst, file=sys.stderr)
        return 2
    shutil.copy2(src, dst)
    print("cloned", src.name, "->", dst)
    return 0


def fill_cover(doc, meta: dict):
    t = doc.Tables(1)
    mapping = (
        (1, 2, meta.get("title", "")),
        (2, 2, meta.get("title_cont", "")),
        (3, 2, meta.get("college", "")),
        (4, 2, meta.get("major", "")),
        (5, 2, meta.get("name", "")),
        (6, 2, meta.get("advisor", "")),
    )
    for r, c, text in mapping:
        replace_cell_text_keep_format(t.Cell(r, c), text)
    for r in (1, 2):
        t.Cell(r, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    date = meta.get("date") or ""
    if date and "起止时间" not in date:
        date = "起止时间：" + date
    if date:
        for i in range(1, min(doc.Paragraphs.Count, 80) + 1):
            p = doc.Paragraphs(i)
            if "起止时间" in para_text(p):
                rng = p.Range
                rng.MoveEnd(wdCharacter, -1)
                rng.Text = date
                try:
                    rng.Font.NameFarEast = "宋体"
                    rng.Font.NameAscii = "Times New Roman"
                    rng.Font.Size = 15
                    rng.Font.Bold = True
                except Exception:
                    pass
                p.Format.Alignment = wdAlignParagraphCenter
                break
    if doc.InlineShapes.Count >= 2:
        doc.InlineShapes(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        doc.InlineShapes(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter


def run_on_doc(args, fn):
    word = attach_word(visible=not args.hidden)
    doc, owned = open_or_reuse(word, Path(args.doc))
    fn(doc)
    doc.Save()
    if args.hidden and owned:
        doc.Close(False)
        try:
            if word.Documents.Count == 0:
                word.Quit()
        except Exception:
            pass


def cmd_cover(args) -> int:
    if not args.json:
        print("cover needs --json", file=sys.stderr)
        return 2
    meta = load_json(Path(args.json))

    def _go(doc):
        fill_cover(doc, meta)
        print("cover filled")

    run_on_doc(args, _go)
    return 0


def insert_picture(sec, path: Path, caption_zh: str, caption_en: str):
    doc = sec.Range.Document
    append_in_section(sec, "", "正文")
    pos = sec.Range.End - 1
    rng = doc.Range(pos, pos)
    shape = rng.InlineShapes.AddPicture(str(path), False, True)
    try:
        max_w = 12.5 * CM
        if shape.Width > max_w:
            ratio = max_w / shape.Width
            shape.Width = max_w
            shape.Height = shape.Height * ratio
    except Exception:
        pass
    try:
        shape.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        shape.Range.ParagraphFormat.FirstLineIndent = 0
        shape.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
    except Exception:
        pass
    for cap, fe, ascii_name in (
        (caption_zh, "黑体", "Arial"),
        (caption_en, "Times New Roman", "Times New Roman"),
    ):
        if not cap:
            continue
        p = append_in_section(sec, cap, "正文")
        p.Format.Alignment = wdAlignParagraphCenter
        p.Format.FirstLineIndent = 0
        try:
            p.Format.CharacterUnitFirstLineIndent = 0
        except Exception:
            pass
        try:
            p.Range.Font.NameFarEast = fe
            p.Range.Font.NameAscii = ascii_name
            p.Range.Font.Size = 10.5
            p.Range.Font.Bold = False
        except Exception:
            pass
    append_in_section(sec, "", "正文")


def insert_omath(sec, linear: str, number: str):
    shown = f"{linear}          {number}" if number.startswith("(") else f"{linear}          ({number})"
    p = append_in_section(sec, shown, "正文")
    p.Format.Alignment = wdAlignParagraphCenter
    p.Format.FirstLineIndent = 0
    try:
        p.Format.CharacterUnitFirstLineIndent = 0
        p.Format.SpaceBefore = 6
        p.Format.SpaceAfter = 6
    except Exception:
        pass
    try:
        fr = p.Range
        fr.MoveEnd(wdCharacter, -1)
        sec.Range.Document.OMaths.Add(fr)
    except Exception as e:
        print("OMath add fail", number, e)


def append_blocks(sec, blocks: list):
    for block in blocks or []:
        kind = block.get("type") or "para"
        if kind == "picture":
            insert_picture(
                sec,
                Path(block["path"]),
                block.get("caption_zh", ""),
                block.get("caption_en", ""),
            )
            continue
        if kind == "omath":
            insert_omath(sec, block["linear"], block.get("number", ""))
            continue
        style = block.get("style") or "正文"
        text = block.get("text", "")
        if isinstance(text, list):
            for line in text:
                append_in_section(sec, line, style)
            continue
        p = append_in_section(sec, text, style)
        if style == "正文" and not text.strip():
            p.Format.FirstLineIndent = 0


def fill_section(doc, section_index: int, payload: dict, *, keep_heading: bool, keep_paras: int):
    sec = doc.Sections(section_index)
    heading = payload.get("heading")
    title_en = payload.get("title_en")
    if title_en:
        set_para_text(sec.Range.Paragraphs(1), title_en)
    elif heading and not keep_heading:
        p1 = sec.Range.Paragraphs(1)
        set_para_text(p1, heading)
        apply_heading1(p1)
    even = payload.get("even_header")
    if even is None and heading and not keep_heading:
        even = heading
    if even:
        set_even_header_text(sec, even)
    if payload.get("fix_odd_header"):
        ensure_odd_header(sec)
    keep_n = int(payload.get("keep_paras") or keep_paras or 1)
    clear_after_para(sec, keep_n)
    append_blocks(sec, payload.get("blocks") or [])
    kw = payload.get("keywords")
    if kw:
        append_in_section(sec, "", "正文")
        p = append_in_section(sec, "关键词：" + kw.removeprefix("关键词："), "正文")
        apply_keywords_zh_format(p)
    kw_en = payload.get("key_words") or payload.get("keywords_en")
    if kw_en:
        append_in_section(sec, "", "正文")
        line = kw_en if kw_en.lower().startswith("key word") else "Key Words：" + kw_en
        p = append_in_section(sec, line, "正文")
        apply_keywords_en_format(p)


def cmd_fill_section(args) -> int:
    payload = load_json(Path(args.json)) if args.json else {}
    if args.heading:
        payload["heading"] = args.heading
    def _go(doc):
        fill_section(
            doc,
            args.section,
            payload,
            keep_heading=args.keep_heading,
            keep_paras=args.keep_paras,
        )
        print("filled section", args.section)

    run_on_doc(args, _go)
    return 0


def extra_chapter(doc, before_index: int, payload: dict):
    heading = payload.get("heading") or ""
    sec = doc.Sections(before_index)
    rng = doc.Range(sec.Range.Start, sec.Range.Start)
    rng.InsertBreak(wdSectionBreakNextPage)
    new_sec = doc.Sections(before_index)
    unlink_headers_keep_footers(new_sec)
    ensure_odd_header(new_sec, ODD_HEADER)
    p1 = new_sec.Range.Paragraphs(1)
    set_para_text(p1, heading)
    apply_heading1(p1)
    set_even_header_text(new_sec, payload.get("even_header") or heading)
    clear_after_para(new_sec, 1)
    append_blocks(new_sec, payload.get("blocks") or [])


def cmd_extra_chapter(args) -> int:
    payload = load_json(Path(args.json)) if args.json else {}
    if args.heading:
        payload["heading"] = args.heading
    def _go(doc):
        extra_chapter(doc, args.before, payload)
        print("inserted chapter before section", args.before, "now sections=", doc.Sections.Count)

    run_on_doc(args, _go)
    return 0


def delete_section(doc, index: int):
    n = doc.Sections.Count
    if index < 1 or index > n:
        raise IndexError(index)
    sec = doc.Sections(index)
    if index == n:
        start = sec.Range.Start
        end = sec.Range.End - 1
        if end > start:
            doc.Range(start, end).Delete()
        return
    sec.Range.Delete()


def cmd_delete_section(args) -> int:
    def _go(doc):
        delete_section(doc, args.index)
        print("deleted section", args.index, "now sections=", doc.Sections.Count)

    run_on_doc(args, _go)
    return 0


def delete_format_tips(doc) -> int:
    n = 0
    for i in range(doc.Paragraphs.Count, 0, -1):
        p = doc.Paragraphs(i)
        t = para_text(p)
        if t.startswith("【格式提示】"):
            p.Range.Delete()
            n += 1
    return n


def apply_all_keyword_fonts(doc) -> int:
    n = 0
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        t = para_text(p)
        if t.startswith("关键词"):
            apply_keywords_zh_format(p)
            n += 1
        elif t.lower().startswith("key word"):
            apply_keywords_en_format(p)
            n += 1
    return n


def fix_master_headers(doc) -> int:
    n = 0
    for i in range(1, doc.Sections.Count + 1):
        sec = doc.Sections(i)
        h = sec.Headers(1)
        rng = h.Range
        rng.MoveEnd(wdCharacter, -1)
        cur = (rng.Text or "").replace("\r", "")
        if "硕士学位论文" in cur or "硕士专业学位" in cur:
            ensure_odd_header(sec, ODD_HEADER)
            n += 1
    return n


def buildup_omaths(doc) -> int:
    n = 0
    for i in range(1, doc.OMaths.Count + 1):
        try:
            doc.OMaths(i).BuildUp()
            n += 1
        except Exception:
            pass
    return n


def cmd_finish(args) -> int:
    def _go(doc):
        tips = delete_format_tips(doc)
        kws = apply_all_keyword_fonts(doc)
        h1 = center_all_heading1(doc)
        hdr = fix_master_headers(doc)
        oms = buildup_omaths(doc)
        try:
            if doc.TablesOfContents.Count:
                doc.TablesOfContents(1).Update()
        except Exception as e:
            print("TOC update fail", e)
        print(
            f"finish: tips={tips} keywords={kws} heading1_centered={h1} "
            f"headers_fixed={hdr} omath_buildup={oms}"
        )

    run_on_doc(args, _go)
    return 0


def add_common(p):
    p.add_argument("--hidden", action="store_true", help="Word Visible=False")
    p.add_argument("--json")
    p.add_argument("--heading")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("clone")
    p.add_argument("--out", required=True)
    p.add_argument("--template")
    p.add_argument("--force", action="store_true")

    p = sub.add_parser("cover")
    add_common(p)
    p.add_argument("--doc", required=True)

    p = sub.add_parser("fill-section")
    add_common(p)
    p.add_argument("--doc", required=True)
    p.add_argument("--section", type=int, required=True)
    p.add_argument("--keep-heading", action="store_true")
    p.add_argument("--keep-paras", type=int, default=1, help="keep first N paras (ABSTRACT=3: 英文题+空行+ABSTRACT)")

    p = sub.add_parser("extra-chapter")
    add_common(p)
    p.add_argument("--doc", required=True)
    p.add_argument("--before", type=int, required=True, help="section index to insert before (结论 is 11 on stock template)")

    p = sub.add_parser("delete-section")
    add_common(p)
    p.add_argument("--doc", required=True)
    p.add_argument("--index", type=int, required=True)

    p = sub.add_parser("finish")
    add_common(p)
    p.add_argument("--doc", required=True)

    args = parser.parse_args(argv)
    handlers = {
        "clone": cmd_clone,
        "cover": cmd_cover,
        "fill-section": cmd_fill_section,
        "extra-chapter": cmd_extra_chapter,
        "delete-section": cmd_delete_section,
        "finish": cmd_finish,
    }
    return handlers[args.cmd](args)


if __name__ == "__main__":
    sys.exit(main())
