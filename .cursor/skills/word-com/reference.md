# Word COM 对象模型参考

集合一律 **从 1 计数**。VBA 里的枚举在 Python 里通常写成整数（见常量表）。

## Application

| 成员 | 说明 |
|------|------|
| `Visible` | `True` 显示窗口 |
| `DisplayAlerts` | `0` 抑制对话框（覆盖确认、转换提示） |
| `ScreenUpdating` | `False` 加快批处理，用户看不见中间帧 |
| `Selection` | 当前选区；自动化里尽量不用，改 `Range` |
| `Documents` | 已打开文档 |
| `NormalTemplate` / `Options` | 一般不要改用户全局选项 |
| `Quit()` | 退出 Word；有未存文档可能弹窗，先 Close |
| `Version` | 如 `"16.0"` |

挂接：

- 已有 Word 窗口：`win32.GetActiveObject("Word.Application")`
- 新开独立实例：`win32.DispatchEx("Word.Application")`
- 避免 `Dispatch("Word.Application")` 绑到坏掉的 `gen_py` 缓存

## Document

打开：`Documents.Open(FileName, ConfirmConversions=False, ReadOnly=False, AddToRecentFiles=False)`

| 成员 | 说明 |
|------|------|
| `FullName` / `Name` / `Path` | 路径 |
| `Paragraphs` `Tables` `Sections` `Styles` | 主集合 |
| `Content` | 正文主故事 Range（不含页眉页脚） |
| `Range(start, end)` | 按字符位置；含节分隔等隐藏字符 |
| `InlineShapes` / `Shapes` | 嵌入图 vs 浮动形状/文本框/WordArt |
| `Comments` `Revisions` `Bookmarks` `Fields` `Hyperlinks` | |
| `OMaths` | 公式对象 |
| `ContentControls` | 内容控件，很多模板没有 |
| `Save()` / `SaveAs2(path, FileFormat=16)` | 16 ≈ `wdFormatDocumentDefault` |
| `Close(SaveChanges)` | `0` 不存 `-1` 存 `2` 提示 |
| `Undo(n)` / `Redo` | Word 原生撤销栈，粒度与 COM 操作不完全对应 |
| `ComputeStatistics(2)` | 页数 `wdStatisticPages=2` |
| `Protect` / `Unprotect` | 表单/修订保护 |

`Content.End - 1` 是最后一段段落标记前。往文末插：

```python
end = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
end.InsertAfter(text + "\r")
```

不要 `Paragraphs.Last.Range.Text = text+"\r"`，容易吃掉分节符。

## Range

自动化的核心。几乎所有改字、改格式都打在 Range 上。

| 成员 | 说明 |
|------|------|
| `Text` | 含段落标记 `\r`；表单元格另含 `\x07` |
| `Start` `End` | 文档绝对偏移 |
| `Font` `ParagraphFormat` `HighlightColorIndex` | |
| `Collapse(0或1)` | `0=wdCollapseEnd` `1=wdCollapseStart` |
| `MoveEnd(Unit, Count)` | `Unit=1` 为 `wdCharacter` |
| `InsertAfter` / `InsertBefore` / `InsertBreak` / `InsertParagraphAfter` | |
| `Copy` `Paste` `Delete` | 剪贴板；跨文档 Copy 可用 |
| `Select` | 仅需给人看/滚动时用 |
| `Find` | 见查找节 |
| `Fields` `Comments` `Bookmarks` | 该范围内的子对象 |
| `Information` | 页码、列等（少用） |

改一段文字且保留段标：

```python
rng = para.Range
rng.MoveEnd(1, -1)  # wdCharacter
rng.Text = "新文本"
```

滚动到可见：`rng.Select()` 后 `word.ActiveWindow.ScrollIntoView(rng, True)`。

## Selection（慎用）

与用户光标是同一个对象。脚本 `TypeText` 会打在光标处，并发编辑会乱。优先构造 `Document.Range` 或 `Paragraphs(i).Range`。

## Font

| 成员 | 说明 |
|------|------|
| `NameFarEast` | 中日韩字体（宋体、黑体、楷体） |
| `NameAscii` | 西文（Times New Roman、Arial） |
| `Name` | 有时只改西文或整段，中英混排不可靠 |
| `NameOther` | **不要设**，pywin32 上常 `0x800a16d4` |
| `Size` | 磅。五号 10.5、小四 12、四号 14、小三 15、三号 16 |
| `Bold` `Italic` | `-1` True，`0` False，`9999999` 混合 |
| `Color` | BGR 整数；自动/未设可能是 `9999999` 或 `-16777216` |
| `Subscript` `Superscript` | |

混合选区的 `Bold`/`Size` 读出来是 `9999999`，表示「不完全一致」。

## Paragraph / ParagraphFormat

| 成员 | 说明 |
|------|------|
| `Style` | 可赋 `doc.Styles("正文")` 或 `NameLocal` |
| `Alignment` | 0 左 1 中 2 右 3 两端 4 分散 |
| `LineSpacingRule` | 0 单倍 1 1.5 倍 2 2 倍 4 固定值 5 多倍 |
| `LineSpacing` | 多倍时：1.25×字号，小四 12 → 15 磅 |
| `SpaceBefore` `SpaceAfter` | 磅 |
| `LineUnitBefore` | 「行」为单位的段前（0.5 行） |
| `FirstLineIndent` | 磅；负数为悬挂 |
| `CharacterUnitFirstLineIndent` | 字符数，中文首行缩进 2 更稳 |
| `LeftIndent` `RightIndent` | |
| `OutlineLevel` | 1–9 大纲，10 正文文本 |
| `DisableLineHeightGrid` | 取消文档网格对齐 |

`Style.NameLocal` 是界面语言下的名字（中文 Word → `标题 1`）。

## Section / PageSetup / 页眉页脚

`InsertBreak(2)` = `wdSectionBreakNextPage`。分页符 `wdPageBreak=7` 不分节，页眉会连到上一节。

`PageSetup`：`PageWidth/Height`、`Top/Bottom/Left/RightMargin`、`HeaderDistance`、`FooterDistance`、`OddAndEvenPagesHeaderFooter`、`DifferentFirstPageHeaderFooter`、`PageNumberStyle`。

页眉页脚索引：

| 常量 | 值 | |
|------|----|----|
| `wdHeaderFooterPrimary` | 1 | 奇数或「默认」 |
| `wdHeaderFooterFirstPage` | 2 | 首页不同时 |
| `wdHeaderFooterEvenPages` | 3 | 偶数页 |

改节页眉前：`header.LinkToPrevious = False`。

页码域：`Fields.Add(range, 33)`（`wdFieldPage`）。重开页码：

```python
pns = section.Footers(1).PageNumbers
pns.RestartNumberingAtSection = True
pns.StartingNumber = 1
```

**不要**赋值 `PageSetup.RestartPageNumbering`（只读）。

`1 cm ≈ 28.35 pt`。A4：595.3 × 841.9 pt。

## Table / Cell

`Tables(i).Cell(row, col)` 1-based。`Cell.Range.Text` = 可见文本 + `\r` + `\x07`。

```python
rng = cell.Range
rng.MoveEnd(1, -1)
rng.Text = "内容"
```

边框（三线表示意）：`Borders(wdBorderTop)` 等；`wdBorderTop=-1` `Bottom=-3` `Left=-2` `Right=-4` `Horizontal=-5`。线宽磅：`LineWidth` 枚举，0.5 磅常用 `wdLineWidth050pt=4`。`LineStyle=1` 单实线。

合并：`cell.Merge(other)`。列宽：`Columns(i).Width`。

## Shapes vs InlineShapes vs Comments

| 类型 | 用途 |
|------|------|
| `InlineShapes` | 嵌入行内的图/OLE |
| `Shapes` | 浮动：文本框、标注、WordArt、自选图形、环绕图 |
| `Comments` | 审阅窗格批注；很多模板 `Count=0` |

读形状文字：`shape.TextFrame.TextRange.Text`（无线条形状会抛错）。倒序 `Delete` 以免索引错乱。`shape.Anchor` 是锚点 Range。

`InlineShapes(i).Type`：1 图，3 常为嵌入对象。宽高单位磅。

## Comments / 审阅

`Comments.Add(Range, Text)`。`Author` / `Initial` 可改。回复：`comment.Replies.Add(Range=comment.Range, Text=...)`（部分版本 `Replies.Count` 仍为 0 时看审阅窗格）。

修订：`doc.TrackRevisions = True` 后再改，然后立刻设回 `False`，避免后续格式操作全进修订。`doc.Revisions.Count` 可探查。高亮：`Range.HighlightColorIndex`（黄=7）。

**润色/最小修改（默认）**：改已有正文一律 `minimal_revision.py`（diff + TrackRevisions），不必等用户说「最小修订」。Agent 每步 **`doc.Save()` 默认执行**；`--no-save` 才跳过。不自动 AcceptAll。

不要改 `Application.UserName`（用户全局选项）。

## Footnotes / Endnotes

```python
mark.Collapse(0)  # wdCollapseEnd
fn = doc.Footnotes.Add(Range=mark, Text="脚注正文")
doc.Endnotes.Add(Range=mark, Text="尾注正文")
# 自定义标记
doc.Footnotes.Add(Range=mark, Text="作者注", Reference="*")
```

给 `fn.Reference` 加书签后，用 `NOTEREF` 域交叉引用（Type=`72`）：`NOTEREF fnFourth \\h`。

脚注与尾注分属两个集合，编号各自独立。

## Captions / 交叉引用

中文 Word 内置标签名常为 `Figure` / `Table` / `Equation`（`CaptionLabels(i).Name` 为准）。`InsertCaption` 打在表格 Range 的 CollapseStart 上：

```python
cap = table.Range
cap.Collapse(1)
cap.InsertCaption(Label="Table", Title="  标题", TitleAutoText="", Position=0)  # 0=上 1=下
```

`REF`（3）/ `PAGEREF`（37）/ `STYLEREF`（48）/ `SEQ`（12）：`Fields.Add` 后写 `Code.Text` 再 `Update()`。不要猜错 Type：`37` 是 PAGEREF 不是引文。

## Bibliography / 引文域

先 `doc.Bibliography.Sources.Add(source_xml)`（OOXML `b:Source`，含唯一 `b:Tag`）。

| 域 | Type | 用法 |
|----|------|------|
| CITATION | **96** | `Fields.Add(rng, 96, "Li2021")` |
| BIBLIOGRAPHY | **97** | `Fields.Add(rng, 97)` |

错误 Type（如 57/75/1980）会插成 SYMBOL/其它域或「数值超出范围」。也可用 Type=`-1`（空域）再设 `Code.Text = " CITATION Li2021 \\l 2052 "`。可选 `Bibliography.BibliographyStyle = "GB7714"`。

## Hyperlinks

```python
doc.Hyperlinks.Add(
    Anchor=rng,
    Address="https://example.com",          # 或 mailto: / 空
    SubAddress="BookmarkName",              # 文档内书签
    TextToDisplay="可见文字",
    ScreenTip="悬停提示",
)
```

`Address=""` + `SubAddress` = 文档内跳转。`HYPERLINK` 域 Type=`88`。

## Fields

`Fields.Add(Range, Type, Text)`。改完 `doc.Fields.Update()`。页眉页脚里的 PAGE 域也算 Fields。

| Type | 值 |
|------|----|
| Empty | -1 |
| REF | 3 |
| XE 索引项 | 4 |
| SEQ | 12 |
| TOC | 13 |
| DATE | 31 |
| PAGE | 33 |
| PAGEREF | 37 |
| STYLEREF | 48 |
| NOTEREF | 72 |
| HYPERLINK | 88 |
| CITATION | 96 |
| BIBLIOGRAPHY | 97 |
| INDEX | 8 |

## Find / Replace

```python
rng = doc.Content
f = rng.Find
f.ClearFormatting()
f.Replacement.ClearFormatting()
f.Text = "旧"
f.Replacement.Text = "新"
f.Forward = True
f.Wrap = 0   # wdFindStop；1=Continue 2=Ask
f.Execute(Replace=1)  # 1 一处 2 全部
```

通配：`f.MatchWildcards = True`（语法与 Word 查找不同，易误伤）。反向替换不能当可靠 Undo。

**注意**：Find/Replace 替换的是**整个 Find 串**；若 LLM 重写整段而 Find 串=整段 old，效果仍像整段替换。润色场景用 `minimal_revision.py` 的 diff 写回。

## Bookmarks

`Bookmarks.Add(name, range)`。名字不能空格；用于定位比裸段落号稳。`Exists(name)` 后 `Bookmarks(name).Range`。

## OMath

```python
rng.MoveEnd(1, -1)
rng.Text = r"E=mc^2"
fr = para.Range
fr.MoveEnd(1, -1)
om = doc.OMaths.Add(fr)
om.BuildUp()
```

线性输入：`_` 下标 `^` 上标 `\sum_` `\int_` `\sqrt` 分数可用 `a/b` 再 BuildUp。编号不要放进线性串。批量删除公式对象易卡死，改文本或整文件快照。

`BuildUp` 失败时：保存后再对已有 `OMaths(i)` 调一次，或保持线性文本。

## 复制跨文档

```python
src.Sections(1).Range.Copy()
dst.Range(0, 0).Paste()
src.Close(False)
```

会带上节属性、页眉、形状。粘贴后检查是否多了后文，按起点删除。

## 常量速查

| 名称 | 值 |
|------|----|
| `wdCharacter` | 1 |
| `wdCollapseEnd` / `Start` | 0 / 1 |
| `wdAlignParagraphLeft/Center/Right/Justify/Distribute` | 0/1/2/3/4 |
| `wdLineSpaceSingle/1pt5/Double/Multiple` | 0/1/2/5 |
| `wdSectionBreakNextPage` | 2 |
| `wdPageBreak` | 7 |
| `wdHeaderFooterPrimary/FirstPage/EvenPages` | 1/2/3 |
| `wdFieldPage` / `PAGEREF` / `CITATION` / `BIBLIOGRAPHY` | 33 / 37 / 96 / 97 |
| `wdPageNumberStyleArabic/UpperRoman` | 0/1 |
| `wdFindStop/Continue` | 0/1 |
| `wdReplaceOne/All` | 1/2 |
| `wdYellow` / `wdNoHighlight` | 7 / 0 |
| `wdFormatDocumentDefault` | 16 |
| `wdStatisticPages` | 2 |
| `wdDoNotSaveChanges` | 0 |

## 稳定性

- 逐段 COM 慢；可见演示才 `sleep`。
- 字体不存在会 COM 错，catch 后降级。
- 异常后杀残留 `WINWORD`，否则文件锁定。
- `gen_py` 损坏时删 `%TEMP%\gen_py` 或 `AppData\Local\Temp\gen_py` 再 DispatchEx。
- 不要并行两个脚本打同一 `.docx`。
