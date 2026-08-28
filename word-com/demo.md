# Word COM 格式演示：复现流程

在 Windows 上用已安装的 Microsoft Word + `pywin32`，按顺序跑脚本，生成可见的演示文档。

| 项 | 值 |
|----|----|
| 产出 | `example/word_com_格式演示.docx` |
| 工作目录 | 仓库根 `wordforcampus` |
| Word 窗口 | 全程保持打开，**不要** `Quit` |
| 同一路径 | **禁止**对已打开的该 `.docx` 再 `Open` 一次（会卡住） |

对象模型与坑见 [reference.md](reference.md)。石大论文排版不是本演示，见 [cupk-thesis-format](../cupk-thesis-format/SKILL.md)。

## 环境

- Windows + Microsoft Word（本机验证为 16.0）
- Python 能 `import win32com.client`
- 从仓库根执行；路径含中文，PowerShell 用引号包住即可

```powershell
cd E:\working\python\python_project\wordforcampus
python -c "import win32com.client; print('pywin32 ok')"
```

## 从零生成（按顺序，各跑一次）

若 Word 里已经打开了 `word_com_格式演示.docx`，先关闭该文档（可留着 Word 进程）。`create_format_demo.py` 会 `Documents.Add()` 再 `SaveAs2` 覆盖该路径；文件被占用时删除/另存会失败。

**不要重复跑第 2–4 步**，它们是往已有文档末尾追加，再跑会内容翻倍。

```powershell
# 1) 新建：字体/颜色/对齐、表格、列表、页眉页脚 PAGE 域、一条 OMath、一个分页符
python .cursor/skills/word-com/scripts/create_format_demo.py

# 2) 追加：删除线/上下标/底纹/边框、编号与符号列表、超链接、批注、三线表、更多 OMath
python .cursor/skills/word-com/scripts/append_format_demo.py

# 3) 追加：脚注/尾注、批注与修订、题注与 REF/PAGEREF/SEQ、文献源与 CITATION/BIBLIOGRAPHY、超链接、TOC、内容控件、首字下沉
python .cursor/skills/word-com/scripts/append_review_refs.py

# 4) 追加：批注回复与高亮修订、更多脚注/NOTEREF、Figure/Table 题注、STYLEREF、更多超链接（含 ScreenTip）、XE/INDEX
python .cursor/skills/word-com/scripts/append_more_review.py

# 5) 收口：合并多余分节，全文只保留「审阅」增补页前的一个分页符；INDEX 改为单栏
python .cursor/skills/word-com/scripts/fix_demo_pagebreaks.py
```

每步结束都会 `Save`，并打印 `Word window left open.` 下一条脚本用 `GetActiveObject` 复用同一窗口、按文件名找到该文档。

## 各步写入什么

```
create_format_demo.py
  标题 1「Word COM 格式能力演示」
  一、字体与颜色　二、表格　三、列表与强调　四、公式
  分页符 → 「第二节…」（同一节，页眉不变）
  书签 DemoStart

append_format_demo.py
  五、更多字符与段落格式　六、三线表　七、更多数学公式

append_review_refs.py
  八、脚注与尾注　九、审阅　十、题注与交叉引用
  十一、文献　十二、超链接　十三、目录　十四、内容控件　十五、首字下沉
  书签 SecReview；Bibliography.Sources：Li2021 / Wang2022

append_more_review.py
  十六、审阅加深　十七、更多脚注　十八、题注　十九、引文域
  二十、超链接　二十一、索引
  书签 MoreReview

fix_demo_pagebreaks.py
  节数 → 1；分页符 → 仅 SecReview 标题前 1 个
  书签 SecReview 不含分页符（否则 REF 域会把分页符复制进正文）
```

## 复现后如何看

Word 里打开产出文件后：

1. **审阅 → 显示批注 / 审阅窗格**：批注、回复、插入/删除修订
2. 页面底部 **脚注**、文末 **尾注**；Ctrl+单击 `NOTEREF` / 超链接 / 书签
3. 悬停带 ScreenTip 的链接
4. 引文域应显示类似 `(李, 2021)`；书目为 `BIBLIOGRAPHY` 域（Type **96 / 97**）
5. 公式若仍是线性文本：公式上右键 → 专业

预期量级（随 Word 版本略有出入）：约 **1 节**、**1 个分页符**、十余页（公式/目录自然占页，不是空页）、批注数条、脚注/尾注、超链接十余个、OMath 十余条、文献源 2 条。

## 只修分页、不重建

文档已在、只是又出现多分节/空白页时，只跑第 5 步。它会先清掉所有分页符，再在「增补页：审阅、脚注、引用与导航」前插回一个。

## 失败时

| 现象 | 处理 |
|------|------|
| 脚本卡住、Word 无响应 | 同一 `.docx` 被 `Open` 了两次。结束多余 `WINWORD`，只留一个窗口再跑 |
| `SaveAs2` / 删除失败 | 先在 Word 里关闭该演示文档 |
| 引文域「数值超出范围」 | CITATION=**96**，BIBLIOGRAPHY=**97**，不要用 37/57/75 |
| 大片空白、短内容独占一页 | 用了「分节符（下一页）」。跑 `fix_demo_pagebreaks.py`；演示脚本应只用 `InsertBreak(7)` |
| `gen_py` / 奇怪 COM 错 | 用 `DispatchEx` / `GetActiveObject`，不要 `Dispatch("Word.Application")`；可删 `%TEMP%\gen_py` |
| 追加脚本跑了两遍 | 内容重复。关掉文档后从第 1 步重来 |

探查（只读统计，仍会挂接已打开文档）：

```powershell
python .cursor/skills/word-com/scripts/inspect_doc.py "example/word_com_格式演示.docx"
python .cursor/skills/word-com/scripts/com_session.py "example/word_com_格式演示.docx"
```
