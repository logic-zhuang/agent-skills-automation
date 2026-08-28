# -*- coding: utf-8 -*-
"""Append review, footnotes, citations, hyperlinks, TOC, captions to the COM demo.

Reuses the open Word document. Does not Quit.
"""
from __future__ import annotations

from pathlib import Path

import win32com.client as win32

wdCharacter = 1
wdCollapseStart = 1
wdCollapseEnd = 0
wdAlignCenter = 1
wdAlignJustify = 3
wdSectionBreakNextPage = 2
wdHeaderFooterPrimary = 1
wdFieldRef = 3
wdFieldPageRef = 37
wdFieldDate = 31
wdFieldSeq = 12
wdFieldHyperlink = 88
wdFieldCitation = 96
wdFieldBibliography = 97
wdFieldEmpty = -1
wdFieldTOC = 13
wdFieldIndex = 8
wdCaptionPositionAbove = 0
wdCaptionPositionBelow = 1
wdContentControlRichText = 0
wdContentControlDate = 6
wdContentControlText = 1
wdGoToBookmark = -1
wdStatisticPages = 2
wdLineStyleSingle = 1
wdLineWidth050pt = 4
wdBorderTop = -1
wdBorderBottom = -3
wdLineWidth150pt = 12

OUT = Path(__file__).resolve().parents[4] / "example" / "word_com_格式演示.docx"


def rgb(r, g, b):
    return r + (g << 8) + (b << 16)


def attach_word():
    try:
        word = win32.GetActiveObject("Word.Application")
    except Exception:
        word = win32.DispatchEx("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    try:
        word.ScreenUpdating = True
    except Exception:
        pass
    return word


def get_doc(word):
    target = str(OUT.resolve()).lower()
    for i in range(1, word.Documents.Count + 1):
        d = word.Documents(i)
        try:
            full = str(Path(d.FullName).resolve()).lower()
        except Exception:
            continue
        if full == target or "格式演示" in d.Name:
            return d
    if not OUT.exists():
        raise FileNotFoundError(OUT)
    return word.Documents.Open(str(OUT), ReadOnly=False)


def add_para(doc, text: str):
    end = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    end.InsertAfter(text + "\r")
    for i in range(doc.Paragraphs.Count, 0, -1):
        cand = doc.Paragraphs(i)
        ct = (cand.Range.Text or "").replace("\r", "").replace("\x07", "")
        if ct == text:
            return cand
    return doc.Paragraphs(doc.Paragraphs.Count - 1)


def body_of(p):
    rng = p.Range
    rng.MoveEnd(wdCharacter, -1)
    return rng


def heading(p, color=None):
    try:
        p.Style = p.Range.Document.Styles("标题 2")
    except Exception:
        pass
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.Size = 16
    r.Font.Bold = True
    r.Font.Color = color or rgb(0, 82, 147)
    p.Format.SpaceBefore = 12


def insert_field(doc, after_para, code: str):
    """Insert a field after paragraph; set Field.Code."""
    rng = after_para.Range
    rng.Collapse(wdCollapseEnd)
    fld = doc.Fields.Add(rng, Type=wdFieldRef)  # placeholder type, overwrite code
    try:
        fld.Code.Text = " " + code + " "
        fld.Update()
    except Exception as e:
        print("field code fail", code, e)
    return fld


def main():
    word = attach_word()
    doc = get_doc(word)
    try:
        word.Activate()
        doc.Activate()
    except Exception:
        pass

    end = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    end.InsertBreak(7)  # wdPageBreak; avoid next-page section (blank remainder pages)

    p = add_para(doc, "增补页：审阅、脚注、引用与导航")
    try:
        p.Style = doc.Styles("标题 1")
    except Exception:
        pass
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.Size = 18
    r.Font.Color = rgb(0, 82, 147)
    p.Format.Alignment = wdAlignCenter
    try:
        doc.Bookmarks.Add("SecReview", body_of(p))
    except Exception:
        pass

    # date field
    p = add_para(doc, "本文档生成/更新日期域：")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    rng = p.Range
    rng.Collapse(wdCollapseEnd)
    rng.MoveStart(wdCharacter, -1)
    # insert date before paragraph mark
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        fld = doc.Fields.Add(ins, wdFieldDate)
        fld.Update()
    except Exception as e:
        print("date field skip", e)

    # ---------- footnotes / endnotes ----------
    p = add_para(doc, "八、脚注与尾注")
    heading(p)

    p = add_para(doc, "脚注用于页面底部的说明。本句末将插入脚注。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    mark = r
    mark.Collapse(wdCollapseEnd)
    try:
        doc.Footnotes.Add(
            Range=mark,
            Text="这是第一条脚注正文，出现在页面底端。COM：Document.Footnotes.Add。",
        )
        print("footnotes", doc.Footnotes.Count)
    except Exception as e:
        print("footnote skip", e)

    p = add_para(doc, "第二条脚注可引用规范或数据来源。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    mark = body_of(p)
    mark.Collapse(wdCollapseEnd)
    try:
        doc.Footnotes.Add(Range=mark, Text="第二条脚注：与第一条自动连续编号。")
    except Exception as e:
        print("footnote2 skip", e)

    p = add_para(doc, "尾注汇总在文末或节末。本句末插入一条尾注。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    mark = body_of(p)
    mark.Collapse(wdCollapseEnd)
    try:
        doc.Endnotes.Add(
            Range=mark,
            Text="这是尾注：Document.Endnotes.Add。通常放在文档或节的末尾。",
        )
        print("endnotes", doc.Endnotes.Count)
    except Exception as e:
        print("endnote skip", e)

    # ---------- review ----------
    p = add_para(doc, "九、审阅：批注与修订")
    heading(p, rgb(192, 0, 0))

    p = add_para(doc, "批注一：对本段添加审阅批注，作者为审阅演示。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    try:
        c1 = doc.Comments.Add(Range=r, Text="请核对此处表述是否与数据一致。")
        c1.Author = "审阅人甲"
        c1.Initial = "甲"
    except Exception as e:
        print("comment1 skip", e)

    p = add_para(doc, "批注二：可对关键词加批注。下面「关键数据」一词带批注。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    f = r.Find
    f.ClearFormatting()
    if f.Execute(FindText="关键数据", Forward=True, Wrap=0):
        try:
            c2 = doc.Comments.Add(Range=r, Text="建议改为具体数值，并注明出处。")
            c2.Author = "审阅人乙"
            c2.Initial = "乙"
        except Exception as e:
            print("comment2 skip", e)

    p = add_para(doc, "修订：下一句将在「修订模式」下被改写，Word 中可见插入/删除标记。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12

    p = add_para(doc, "这是修订前的原始句子，准备被替换。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    try:
        doc.TrackRevisions = True
    except Exception:
        pass
    try:
        rr = body_of(p)
        rr.Text = "这是修订后的句子（TrackRevisions=True 时应显示为修订）。"
    except Exception as e:
        print("revision edit skip", e)
    try:
        doc.TrackRevisions = False  # stop tracking further script edits
    except Exception:
        pass
    print("revisions", doc.Revisions.Count, "comments", doc.Comments.Count)

    p = add_para(doc, "审阅窗格可查看批注与修订。本脚本关闭后续修订跟踪，以免把格式操作都记成修订。")
    r = body_of(p)
    r.Font.NameFarEast = "楷体"
    r.Font.Size = 10.5
    r.Font.Color = rgb(89, 89, 89)

    # ---------- captions / cross-ref / seq ----------
    p = add_para(doc, "十、题注、题注域与交叉引用")
    heading(p)
    try:
        doc.Bookmarks.Add("CaptionSection", body_of(p))
    except Exception:
        pass

    p = add_para(doc, "先插入一个示意表格，再加题注。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12

    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    table = doc.Tables.Add(rng, 2, 2)
    for i, row in enumerate((("A", "B"), ("1", "2")), 1):
        for c, val in enumerate(row, 1):
            cell = table.Cell(i, c)
            cr = cell.Range
            cr.MoveEnd(wdCharacter, -1)
            cr.Text = val
            cr.Font.Size = 10.5
            cell.Range.ParagraphFormat.Alignment = wdAlignCenter
    try:
        table.Borders.Enable = True
    except Exception:
        pass
    table.Rows.Alignment = 1

    try:
        cap_rng = table.Range
        cap_rng.Collapse(wdCollapseStart)
        cap_rng.InsertCaption(
            Label="Table",
            Title="  题注演示表（InsertCaption）",
            TitleAutoText="",
            Position=wdCaptionPositionAbove,
        )
        print("caption inserted")
    except Exception as e:
        print("caption skip", e)
        p = add_para(doc, "表 SEQ  题注演示表（InsertCaption 失败时的后备文本）")
        body_of(p).Font.NameFarEast = "黑体"
        body_of(p).Font.Size = 10.5
        p.Format.Alignment = wdAlignCenter

    p = add_para(doc, "交叉引用：请参见书签「增补页标题」（REF 域，可 Ctrl+单击跳转）。此处插入 PAGE 交叉页码。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2

    # REF field to SecReview bookmark
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        fld = doc.Fields.Add(ins, wdFieldRef)
        fld.Code.Text = " REF SecReview \\h "
        fld.Update()
    except Exception as e:
        print("REF skip", e)

    p = add_para(doc, "PAGEREF：标题所在页为第 ")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        fld = doc.Fields.Add(ins, wdFieldPageRef)
        fld.Code.Text = " PAGEREF SecReview \\h "
        fld.Update()
    except Exception as e:
        print("PAGEREF skip", e)
    # add 页 after field - insert after paragraph
    p2 = add_para(doc, " 页（PAGEREF 域）。SEQ 可用于公式/图表自动编号。")
    body_of(p2).Font.NameFarEast = "宋体"
    body_of(p2).Font.Size = 12

    p = add_para(doc, "当前 SEQ 图编号：")
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        fld = doc.Fields.Add(ins, wdFieldSeq)
        fld.Code.Text = " SEQ 图 "
        fld.Update()
    except Exception as e:
        print("SEQ skip", e)

    # ---------- bibliography / citation ----------
    p = add_para(doc, "十一、文献源、引文域与书目")
    heading(p)

    p = add_para(doc, "Word 内置书目：先向 Bibliography.Sources 添加 Source XML，再插入 CITATION / BIBLIOGRAPHY 域。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2

    src_xml = """<b:Source xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">
  <b:Tag>Li2021</b:Tag>
  <b:SourceType>Book</b:SourceType>
  <b:Author><b:Author><b:NameList>
    <b:Person><b:Last>李</b:Last><b:First>四</b:First></b:Person>
  </b:NameList></b:Author></b:Author>
  <b:Title>COM 自动化与科技论文排版</b:Title>
  <b:Year>2021</b:Year>
  <b:City>北京</b:City>
  <b:Publisher>示例出版社</b:Publisher>
</b:Source>"""
    src_xml2 = """<b:Source xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">
  <b:Tag>Wang2022</b:Tag>
  <b:SourceType>ArticleInAPeriodical</b:SourceType>
  <b:Author><b:Author><b:NameList>
    <b:Person><b:Last>王</b:Last><b:First>五</b:First></b:Person>
  </b:NameList></b:Author></b:Author>
  <b:Title>脚注与交叉引用实践</b:Title>
  <b:Year>2022</b:Year>
  <b:PeriodicalTitle>演示学报</b:PeriodicalTitle>
  <b:Volume>10</b:Volume>
  <b:Issue>2</b:Issue>
  <b:Pages>12-18</b:Pages>
</b:Source>"""
    for xml in (src_xml, src_xml2):
        try:
            doc.Bibliography.Sources.Add(xml)
        except Exception as e:
            print("source add skip", e)

    p = add_para(doc, "正文引用示例：有学者讨论了自动化排版问题（引文域后接）。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        fld = doc.Fields.Add(ins, wdFieldCitation, "Li2021")
        fld.Update()
    except Exception as e:
        print("citation1 skip", e)

    p = add_para(doc, "期刊文献的引用：")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        fld = doc.Fields.Add(ins, wdFieldCitation, "Wang2022")
        fld.Update()
    except Exception as e:
        print("citation2 skip", e)

    p = add_para(doc, "参考文献（BIBLIOGRAPHY 域，样式随 Word 书目样式变化）")
    try:
        p.Style = doc.Styles("标题 3")
    except Exception:
        pass
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.Size = 14

    p = add_para(doc, "")
    ins = body_of(p)
    try:
        try:
            doc.Bibliography.BibliographyStyle = "GB7714"
        except Exception:
            try:
                doc.Bibliography.BibliographyStyle = "APA"
            except Exception:
                pass
        fld = doc.Fields.Add(ins, wdFieldBibliography)
        fld.Update()
        print("bibliography field ok, sources", doc.Bibliography.Sources.Count)
    except Exception as e:
        print("bibliography skip", e)
        add_para(doc, "[1] 李四. COM 自动化与科技论文排版[M]. 北京: 示例出版社, 2021.")
        add_para(doc, "[2] 王五. 脚注与交叉引用实践[J]. 演示学报, 2022, 10(2): 12-18.")

    # ---------- hyperlinks ----------
    p = add_para(doc, "十二、更多超链接")
    heading(p, rgb(0, 100, 0))

    p = add_para(doc, "外部网址：打开 Word VBA 文档。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    try:
        doc.Hyperlinks.Add(
            Anchor=r,
            Address="https://learn.microsoft.com/office/vba/api/word.footnotes",
            TextToDisplay="Microsoft 文档：Footnotes 对象",
        )
    except Exception as e:
        print("hl url skip", e)

    p = add_para(doc, "邮箱链接：发送邮件。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    try:
        doc.Hyperlinks.Add(
            Anchor=r,
            Address="mailto:reviewer@example.com?subject=Word%20COM%20Demo",
            TextToDisplay="reviewer@example.com",
        )
    except Exception as e:
        print("hl mail skip", e)

    p = add_para(doc, "文档内链接：回到本页标题书签 SecReview。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    try:
        doc.Hyperlinks.Add(
            Anchor=r,
            Address="",
            SubAddress="SecReview",
            TextToDisplay="跳转到「增补页：审阅、脚注、引用与导航」",
        )
    except Exception as e:
        print("hl bookmark skip", e)

    p = add_para(doc, "若存在 DemoStart 书签（文首），也可跳转：")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    try:
        if doc.Bookmarks.Exists("DemoStart"):
            doc.Hyperlinks.Add(
                Anchor=r,
                Address="",
                SubAddress="DemoStart",
                TextToDisplay="跳转到文档开头标题",
            )
    except Exception as e:
        print("hl demostart skip", e)

    # ---------- TOC ----------
    p = add_para(doc, "十三、目录域（标题样式）")
    heading(p)
    p = add_para(doc, "下面用 TablesOfContents.Add 基于「标题 1/2」生成目录，更新域后可见页码。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2

    p = add_para(doc, "")
    try:
        toc_rng = body_of(p)
        doc.TablesOfContents.Add(
            Range=toc_rng,
            UseHeadingStyles=True,
            UpperHeadingLevel=1,
            LowerHeadingLevel=2,
            UseHyperlinks=True,
        )
        print("TOC count", doc.TablesOfContents.Count)
        if doc.TablesOfContents.Count:
            doc.TablesOfContents(1).Update()
    except Exception as e:
        print("TOC Add skip, try field", e)
        try:
            ins = body_of(p)
            fld = doc.Fields.Add(ins, wdFieldTOC)
            fld.Code.Text = " TOC \\o \"1-2\" \\h \\z \\u "
            fld.Update()
        except Exception as e2:
            print("TOC field skip", e2)

    # ---------- content controls ----------
    p = add_para(doc, "十四、内容控件与下拉（可选）")
    heading(p)
    p = add_para(doc, "日期控件占位：")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        cc = doc.ContentControls.Add(wdContentControlDate, ins)
        cc.Title = "填写日期"
        cc.DateDisplayFormat = "yyyy年M月d日"
        print("content controls", doc.ContentControls.Count)
    except Exception as e:
        print("content control skip", e)

    p = add_para(doc, "纯文本内容控件：")
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        cc = doc.ContentControls.Add(wdContentControlText, ins)
        cc.Title = "姓名"
        cc.Range.Text = "（点击填写）"
    except Exception as e:
        print("cc text skip", e)

    # drop cap on a new para
    p = add_para(doc, "十五、首字下沉")
    heading(p)
    p = add_para(doc, "首字下沉把段落第一个字放大占两到三行，常用于杂志开篇。COM：Paragraph.DropCap。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 0
    try:
        p.DropCap.Enable()
        p.DropCap.LinesToDrop = 2
        p.DropCap.FontName = "黑体"
        print("dropcap ok")
    except Exception as e:
        print("dropcap skip", e)

    # two-column snippet in a nested... changing whole section columns is aggressive.
    # Use a frame or skip. Optional: textboxes side by side already have one shape.

    try:
        doc.Fields.Update()
    except Exception:
        pass
    try:
        if doc.TablesOfContents.Count:
            doc.TablesOfContents(1).Update()
    except Exception:
        pass

    try:
        doc.Save()
    except Exception:
        doc.SaveAs2(str(OUT), FileFormat=16)

    print(
        "saved",
        OUT,
        "sections",
        doc.Sections.Count,
        "comments",
        doc.Comments.Count,
        "revisions",
        doc.Revisions.Count,
        "footnotes",
        doc.Footnotes.Count,
        "endnotes",
        doc.Endnotes.Count,
        "hyperlinks",
        doc.Hyperlinks.Count,
        "fields",
        doc.Fields.Count,
        "toc",
        doc.TablesOfContents.Count,
        "cc",
        doc.ContentControls.Count,
    )
    print("Word window left open.")


if __name__ == "__main__":
    main()
