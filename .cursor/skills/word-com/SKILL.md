---
name: Word文档操作
description: >-
  Windows Word COM 执行层：打开/复用文档、Markdown→Word（md_to_docx）、inspect、
  最小修订（TrackRevisions+diff）、填表与混合文档。不含学校论文格式；须与论文套版
  技能同时加载。用户提到 Word、Markdown 转 Word、润色、批改、改稿、填表、修订、
  track changes 时使用。
source: repo
---

# Word文档操作（执行层）

Windows + Microsoft Word + `pywin32`。**只负责怎么动 Word**，不含学校/院系格式规则。  
与 [wps-com](../wps-com/SKILL.md) **产品能力对齐**（MD→docx / 润色 / 填表 / hybrid）；本 skill 仅挂接 Microsoft Word。

论文套版（模板、节映射、页眉页码规则）由 **`cupk-thesis-format` 等 `*-thesis-format` skill** 定义；本 skill 提供 COM 执行约定与通用场景脚本。

## 与学校 thesis-format skill 配合（必守）

宿主 App / Agent 会话应 **同时加载**：

```text
["Word文档操作", "CUPK本科论文套版", ...]
```

| 任务 | 读哪个 skill | 跑哪个脚本 |
|------|--------------|------------|
| **Markdown → Word**（OpenWorker 写稿） | 本 skill → [scenarios/md-to-docx.md](scenarios/md-to-docx.md) | `scripts/md_to_docx.py` |
| 润色 / 批改 / 改字改段 | 本 skill → [scenarios/revise.md](scenarios/revise.md) | `scripts/minimal_revision.py` |
| 填固定表格 | 本 skill → [scenarios/form-fill.md](scenarios/form-fill.md) | `inspect` → Agent `fill.json` → `auto_fill_form --fill --draft` |
| 表 + 分段框架（申请书等） | 本 skill → [scenarios/hybrid.md](scenarios/hybrid.md) | `scripts/hybrid_ops.py` |
| 某校论文套版 | **学校 skill** → `convert.md` / `rules.md` | 学校 `scripts/convert_ops.py` 等 |

学校 skill 内所有 COM 操作 **遵守本 skill 硬规则**（Open 复用、Range、TrackRevisions、Save、勿重写 PAGE 域等）。

OpenWorker：`load_skill("Word文档操作")` 与 `load_skill("CUPK本科论文套版")` 各得 `resources_path`；**用户来稿只读**；填写/修订输出写在**会话工作区**。技能目录名仍为 `word-com` / `cupk-thesis-format`。

**访问规则（必守）**：shell cwd 保持 OpenWorker 会话工作区；用 `"<resources_path>/scripts/..."`（路径用 `/`）跑脚本。`load_skill` 后可用 `list_files` / `read_file` 读技能目录；若仍报「超出访问范围」，改用 shell（`dir` / `Get-ChildItem`），**不要** `cd` 到 `resources_path`。

## 填表分工（Agent 主导）

| 层 | 负责 |
|----|------|
| `inspect_form` | `facts` + 可选 `suggestions` |
| Agent | 判版式、写/改 `fill.json`；`--auto-suggest` 仅旧表回归 |
| `auto_fill_form --fill` / `fill_form` | clone→draft、按计划写入 |

详见 [scenarios/form-fill.md](scenarios/form-fill.md)。

## 稳定反馈契约（必守）

体系说明见仓库 [SKILLS_SYSTEM.md](../SKILLS_SYSTEM.md)。本 skill 执行时：

1. **入口**：用户说 `/word`、Word、docx → 用本 skill；说 WPS/金山 → **转** [wps-com](../wps-com/SKILL.md)，不静默混用。
2. **场景表命中即停**（下一节）；缺文件路径 / 缺「改哪一段」→ 先问再跑脚本。
3. **禁跳步**：填表必须 inspect → fill → `--draft`；润色必须 `minimal_revision`（禁止整段覆盖）；原文件只读。
4. **回复骨架**（每次）：判定场景 → 已做（脚本）→ 产物路径 → 待确认 → 下一步。失败同样结构，写明停点。
5. **不** AcceptAll；填表默认不定稿，确认后再 `--save`。

## 场景判定（Agent 先看用户意图）

按顺序匹配，命中即停：

1. **Markdown 转 Word / 只有 .md 要出 docx** → [scenarios/md-to-docx.md](scenarios/md-to-docx.md)
2. **改已有正文**（润色、批改、改几个字、改第 N 段）→ [scenarios/revise.md](scenarios/revise.md)
3. **只填表格字段**（导员表、几项 cover 式字段）→ [scenarios/form-fill.md](scenarios/form-fill.md)
4. **论文套版 / 某校格式 / 毕设排版**→ **转已加载的 `*-thesis-format` skill**（未加载则提示安装对应学校包；来稿若是 MD，先 `md_to_docx`）
5. **上半表格 + 下面分部分正文**（申请书）→ [scenarios/hybrid.md](scenarios/hybrid.md)
6. **不确定** → `scripts/inspect_doc.py` 探查后再判

易混：

- 「把 md 转成 word」→ **md-to-docx**，不是套版
- 「改第 2 章表述」→ **批改**（revise），不是套版
- 「套成石大本科格式」→ 若只有 MD：先 md-to-docx，再 **学校 skill**（cupk）

## 必读

| 文档 / 脚本 | 用途 |
|-------------|------|
| [reference.md](reference.md) | COM 对象、常量、坑 |
| [minimal-revision.md](minimal-revision.md) | 最小修订策略详解 |
| [scenarios/md-to-docx.md](scenarios/md-to-docx.md) | OpenWorker MD → 粗 docx |
| [scenarios/revise.md](scenarios/revise.md) | 批改流程 |
| [scenarios/form-fill.md](scenarios/form-fill.md) | 填表：结构报告 → plan → 修订草稿 → 用户确认再定稿 |
| [kb-profile.md](kb-profile.md) | 知识库 L1/L2；用户不写 JSON |
| [scenarios/hybrid.md](scenarios/hybrid.md) | 混合文档流程 |
| [scripts/md_to_docx.py](scripts/md_to_docx.py) | Markdown → .docx（python-docx） |
| [scripts/auto_fill_form.py](scripts/auto_fill_form.py) | 填表流水线（默认不定稿 + TrackRevisions） |
| [scripts/inspect_form.py](scripts/inspect_form.py) | 结构报告 v2 |
| [scripts/plan_fill.py](scripts/plan_fill.py) | slots + profile → fill；needs_confirm |
| [scripts/parse_profile.py](scripts/parse_profile.py) | 自然语言/键值 → 内部 profile |
| [scripts/fill_form.py](scripts/fill_form.py) | COM 写入 cells |
| [scripts/form_utils.py](scripts/form_utils.py) | 结构探测 / 别名 / verify |
| [scripts/minimal_revision.py](scripts/minimal_revision.py) | 最小 diff 修订写回 |
| [scripts/inspect_doc.py](scripts/inspect_doc.py) | 只读探查 |
| [scripts/com_session.py](scripts/com_session.py) | 打开并保持 Word 窗口 |

## 默认改稿策略（改已有正文）

不必等用户说「最小修订」：

- 用 `minimal_revision.py` + `TrackRevisions=True` + `SequenceMatcher` diff
- 默认 **Save**（`--no-save` 才跳过）
- **不**自动 `AcceptAll`
- **禁止**整段 `Paragraph.Range.Text = 新全文` / `doc.Content.Text = 新全文`

例外（可整段 `Range.Text`）：新建空白段、学校 skill 的 `fill-section`、用户明确不要修订痕迹。

```bash
python <resources_path>/scripts/minimal_revision.py "doc.docx" --paragraph 2 --new "润色后段落"
python <resources_path>/scripts/minimal_revision.py "doc.docx" --replace "旧" "新"
python <resources_path>/scripts/minimal_revision.py --active --replace "旧" "新"
```

## 工作流程

```
- [ ] 确认 Word + pywin32；MD 转 Word 另需 python-docx
- [ ] 仅有 Markdown → md_to_docx.py → 会话目录 .docx，再进入后续场景
- [ ] GetActiveObject 复用；禁止同路径二次 Open
- [ ] DisplayAlerts=0；Visible=True
- [ ] 填表 → parse_profile（若需）→ inspect → plan → auto_fill（--draft，勿默认 --save）→ 用户审阅确认后再定稿
- [ ] 批改 → minimal_revision（默认 Save）；填表默认不定稿；均禁止 AcceptAll
- [ ] 不用 Selection 改内容
```

## COM 硬规则

1. COM 集合 **1-based**；段末 `\r`，单元格 `\x07`。
2. **改已有正文** → `minimal_revision.py`。**新建/学校 fill-section** → `MoveEnd(1,-1)` 再 `Text = …`。
3. 中英混排：`Font.NameFarEast` + `Font.NameAscii`；不设 `NameOther`。
4. 样式名带空格：`"标题 1"`、`"正文"`。
5. 页码：`Footers(i).PageNumbers.RestartNumberingAtSection`；不写 `PageSetup.RestartPageNumbering`。
6. 奇偶页眉：先 `LinkToPrevious = False`。
7. 分节：`InsertBreak(2)`，不要只用分页符。
8. OMath：`OMaths.Add` → `BuildUp()`；勿批量 `OMaths.Range.Delete`。
9. 跳过 `~$*.docx`；同文件勿二次 Open。

## 探查

```bash
python <resources_path>/scripts/inspect_doc.py "path/to.docx"
python <resources_path>/scripts/com_session.py "path/to.docx"
```

## 与 python-docx

- **Markdown → 粗 docx**：用 `scripts/md_to_docx.py`（不必开 Word）。
- 页眉奇偶、分节、域、OMath、TrackRevisions、可见窗口、学校套版 → **必须 COM**（及对应学校 skill）。
