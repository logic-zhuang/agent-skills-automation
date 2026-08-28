# -*- coding: utf-8 -*-
"""Create a minimal undergraduate thesis .docx for CUPK Phase 3 testing."""
from __future__ import annotations

from pathlib import Path

import win32com.client as win32

wdCharacter = 1
wdCollapseEnd = 0
wdAlignParagraphCenter = 1
wdSectionBreakNextPage = 2

OUT = Path(__file__).resolve().parent / "thesis_source.docx"


def attach_word():
    try:
        word = win32.GetActiveObject("Word.Application")
    except Exception:
        word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    return word


def set_text(p, text: str):
    rng = p.Range
    rng.MoveEnd(wdCharacter, -1)
    rng.Text = text


def add_para(doc, text: str, style: str = "正文"):
    p = doc.Paragraphs.Add()
    set_text(p, text)
    try:
        p.Style = style
    except Exception:
        pass
    return p


def add_h1(doc, text: str):
    p = add_para(doc, text, "标题 1")
    p.Format.Alignment = wdAlignParagraphCenter
    return p


def add_h2(doc, text: str):
    return add_para(doc, text, "标题 2")


def main() -> int:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    if OUT.exists():
        OUT.unlink()

    word = attach_word()
    doc = word.Documents.Add()
    try:
        t = doc.Tables.Add(doc.Range(0, 0), 6, 2)
        rows = (
            ("题目", "Phase3 套版测试论文"),
            ("", ""),
            ("学院名称", "石油学院"),
            ("专业名称", "石油工程"),
            ("学生姓名", "测试同学"),
            ("指导教师", "张老师 副教授"),
        )
        for r, (label, value) in enumerate(rows, start=1):
            set_text(t.Cell(r, 1).Range.Paragraphs(1), label)
            set_text(t.Cell(r, 2).Range.Paragraphs(1), value)

        add_para(doc, "起止时间：2025 年 9 月 1 日 至 2026 年 6 月 15 日")
        doc.Range(doc.Range().End - 1, doc.Range().End - 1).InsertBreak(wdSectionBreakNextPage)

        add_h1(doc, "摘    要")
        add_para(
            doc,
            "本文以 Phase 3 流水线为对象，验证 CUPK 本科论文套版脚本能否从来稿提取内容、"
            "克隆模板并按节填充。测试覆盖摘要、英文摘要、两章正文、结论、参考文献与致谢。",
        )
        add_para(doc, "关键词：套版测试；Word COM；本科论文")

        doc.Range(doc.Range().End - 1, doc.Range().End - 1).InsertBreak(wdSectionBreakNextPage)
        add_para(doc, "Phase 3 Pipeline Test Thesis", "标题 1")
        add_para(doc, "")
        add_h1(doc, "ABSTRACT")
        add_para(
            doc,
            "This document exercises the CUPK undergraduate thesis formatting pipeline. "
            "It includes bilingual abstracts, two chapters, a conclusion, references, and acknowledgements.",
        )
        add_para(doc, "Key Words: formatting test; Word COM; undergraduate thesis")

        doc.Range(doc.Range().End - 1, doc.Range().End - 1).InsertBreak(wdSectionBreakNextPage)
        add_h1(doc, "前    言")
        add_para(doc, "本测试来稿用于验证 convert_ops 脚本，不应直接在原稿上套版。")

        add_h1(doc, "第1章  绪论")
        add_h2(doc, "1.1  研究背景")
        add_para(doc, "本科毕业论文排版需要严格遵循校区格式范本，COM 自动化可降低手工出错率。")
        add_para(doc, "k=a\\phi^b/S_wi^c          (1.1)")

        add_h1(doc, "第2章  研究方法")
        add_h2(doc, "2.1  脚本流程")
        add_para(doc, "流程包括 inspect_source、clone、cover、fill-section 与 finish。")

        add_h1(doc, "结    论")
        add_para(doc, "Phase 3 测试来稿结构完整，可用于验证套版脚本。")

        add_h1(doc, "参考文献")
        add_para(doc, "[1] 中国石油大学（北京）克拉玛依校区. 本科毕业设计（论文）格式范本（2023年修订）.")

        add_h1(doc, "致    谢")
        add_para(doc, "感谢指导教师在测试方案设计中的帮助。")

        doc.SaveAs2(str(OUT.resolve()))
        print("wrote", OUT)
    finally:
        doc.Close(False)
        try:
            if word.Documents.Count == 0:
                word.Quit()
        except Exception:
            pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
