---
name: Word文档操作
description: Windows Word COM（主路径：打开 Word 可见改稿）。python-docx 仅 md→粗 docx 桥；写稿/润色/填表/修订一律 COM。不含学校论文格式。提及 Word、润色、批改、填表、修订时使用。
source: repo
---

# Word文档操作（执行层）

Windows + Microsoft Word + `pywin32`。**只负责怎么动 Word**，不含学校/院系格式规则。  
与 [wps-com](../wps-com/SKILL.md) **产品能力对齐**（COM 写稿 / 润色 / 填表 / hybrid）；本 skill 仅挂接 Microsoft Word。

## python-docx 与 Word COM 职责（必守）

本 skill 有**两条技术路径**，分工不同。**禁止在同一条编辑链路上混用**。

| | **python-docx** | **Word COM**（微软 `Word.Application` + pywin32） |
|--|-----------------|--------------------------------------------------|
| **本质** | Python 直接读写 `.docx` 里的 XML | 操控**已安装、正在运行**的 Word |
| **Word 要不要开** | **不要** | **要**（`attach_word`；默认 `Visible=True`） |
| **用户能否即时看到改动** | **不能**；须保存后重新打开 | **能**；已开窗口里实时刷新 |
| **修订痕迹**（删除线/下划线） | **无** | **有**（`TrackRevisions` + `minimal_revision`） |
| **页眉页码、分节、域、OMath** | 弱/无 | **完整** |
| **本 skill 中的脚本** | **仅** `md_to_docx.py` | **其余全部**（author / revise / fill / hybrid / inspect / 学校套版 COM） |

### 何时用哪条

```text
写稿 / 填表 / 润色 / 套版 / 用户在 Word 里审改     →  Word COM（默认）
用户已有 .md，要一个能打开的粗 .docx              →  python-docx（md_to_docx）→ 再用 COM 打开后续
```

### Agent 硬规则

1. **OpenWorker 默认 COM**；不要为写稿先造 `.md` 再 `md_to_docx`（见 [author.md](scenarios/author.md)）。
2. **`md_to_docx` 产出是离线草稿**；润色、填表、套版、可见修订 **必须** 转 COM，**不得** 继续用 python-docx 改同一文件。
3. **同一 docx 一轮任务**：要么全程 COM，要么「仅 md_to_docx 转换一步」后切 COM；**禁止** COM 改完又用 python-docx 回写。
4. 用户要「**打开 Word 就能看到效果**」→ **只能 COM** + `Visible=True`；改已有字还要 **审阅 → 所有标记**。

**OpenWorker 写稿默认：COM 直接生成/写入 `.docx`**（见 [scenarios/author.md](scenarios/author.md)）。`md_to_docx` 仅当用户已有 `.md` 或明确要求转换。

论文套版由 **`cupk-thesis-format` 等 `*-thesis-format` skill** 定义；本 skill 提供 COM 执行约定与通用场景脚本。

## 与学校 thesis-format skill 配合（必守）

宿主 App / Agent 会话应 **同时加载**：

```text
["Word文档操作", "CUPK本科论文套版", ...]
```

| 任务 | 读哪个 skill | 跑哪个脚本 |
|------|--------------|------------|
| **从零写稿 / 新建 docx**（OpenWorker 默认） | 本 skill → [scenarios/author.md](scenarios/author.md) | COM：`attach_word` + `Documents.Add` / 会话脚本 |
| 润色 / 批改 / 改字改段 | 本 skill → [scenarios/revise.md](scenarios/revise.md) | `scripts/minimal_revision.py` |
| 填固定表格 | 本 skill → [scenarios/form-fill.md](scenarios/form-fill.md) | `inspect` → Agent `fill.json` → `auto_fill_form --fill --draft` |
| 表 + 分段框架（申请书等） | 本 skill → [scenarios/hybrid.md](scenarios/hybrid.md) | `scripts/hybrid_ops.py` |
| **已有 MD → docx**（可选，非默认） | 本 skill → [scenarios/md-to-docx.md](scenarios/md-to-docx.md) | `scripts/md_to_docx.py` |
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
3. **禁跳步**：写稿默认 COM 直出 docx，**禁止**无必要先造 md 再转；填表必须 inspect → fill → `--draft`；**任何改已有正文**（含删空格/无用字符）必须 `minimal_revision`；原文件只读。用户须在 Word **所有标记** 下见改前（删除线）与改后（下划线）。
4. **只用 Word COM**：`attach_word` / 本 skill 脚本；**禁止**用系统默认程序打开 docx（易误开 WPS）。`attach_word` 若发现 ROT 里的 `Word.Application` 实为 WPS（Path 含 Kingsoft/WPS），**拒绝复用**并 `DispatchEx` 拉真 Word；若注册表也被劫持则报错并提示转 **wps-com**，禁止假装成功。
5. **多选 □**：只改 `□`→`☑`，禁止删除其他选项（见 form-fill.md）。
6. **回复骨架**（每次）：判定场景 → 已做（脚本）→ 产物路径 → `backend: Word.Application`（以脚本打印的 `path=` 为准，须为 Microsoft Office，不是 Kingsoft）→ 待确认 → 下一步。失败同样结构，写明停点。
7. **不** AcceptAll；填表默认不定稿，确认后再 `--save`。

## 场景判定（Agent 先看用户意图）

按顺序匹配，命中即停：

1. **从零写稿 / 生成 Word / 新建 docx** → [scenarios/author.md](scenarios/author.md)（**默认**；勿先 md）
2. **改已有正文**（润色、批改、改几个字、删无用字符/空格、改第 N 段）→ [scenarios/revise.md](scenarios/revise.md)
3. **只填表格字段**（导员表、几项 cover 式字段）→ [scenarios/form-fill.md](scenarios/form-fill.md)
4. **论文套版 / 某校格式 / 毕设排版**→ **转已加载的 `*-thesis-format` skill**（未加载则提示安装；来稿若是 MD 且用户要套版，再 `md_to_docx`）
5. **上半表格 + 下面分部分正文**（申请书）→ [scenarios/hybrid.md](scenarios/hybrid.md)
6. **用户已有 .md 或明确「md 转 Word」** → [scenarios/md-to-docx.md](scenarios/md-to-docx.md)
7. **不确定** → `scripts/inspect_doc.py` 探查后再判

易混：

- 「写一篇 / 生成 Word」→ **author（COM）**，不是 md-to-docx
- 「把 md 转成 word」→ **md-to-docx**，不是套版
- 「改第 2 章表述」→ **批改**（revise），不是套版
- 「套成石大本科格式」→ 来稿已是 Word 则直接学校 skill；仅有 MD 时才先 md-to-docx

## 必读

| 文档 / 脚本 | 用途 |
|-------------|------|
| [reference.md](reference.md) | COM 对象、常量、坑；python-docx 边界见 SKILL「职责」节 |
| [minimal-revision.md](minimal-revision.md) | 最小修订策略详解 |
| [scenarios/author.md](scenarios/author.md) | OpenWorker 默认 COM 写稿 |
| [scenarios/revise.md](scenarios/revise.md) | 批改流程 |
| [scenarios/form-fill.md](scenarios/form-fill.md) | 填表：结构报告 → plan → 修订草稿 → 用户确认再定稿 |
| [scenarios/md-to-docx.md](scenarios/md-to-docx.md) | 可选：已有 MD → 粗 docx |
| [kb-profile.md](kb-profile.md) | 知识库 L1/L2；用户不写 JSON |
| [scenarios/hybrid.md](scenarios/hybrid.md) | 混合文档流程 |
| [scripts/md_to_docx.py](scripts/md_to_docx.py) | Markdown → .docx（python-docx；非默认） |
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

不必等用户说「最小修订」。**OpenWorker 对正文的任何修改都必须留下可见修订**（含「删无用字符」）：

- 一律 `minimal_revision.py`：`TrackRevisions=True` + `SequenceMatcher`；脚本会切到 **所有标记**
- 删字符 = `--replace "待删" ""` 或段级 old→new（diff 只 `Delete`），Word 里应见**删除线**，不是字直接消失
- 默认 **Save**（`--no-save` 才跳过）；**不** `AcceptAll`
- **禁止**：整段 `Range.Text = 新全文`、静默 `Delete`/`TypeBackspace`、关修订再改

例外（可无修订痕迹）：新建空白段、学校 skill 的 `fill-section`、用户明确不要修订痕迹。

```bash
python <resources_path>/scripts/minimal_revision.py "doc.docx" --paragraph 2 --new "润色后段落"
python <resources_path>/scripts/minimal_revision.py "doc.docx" --replace "旧" "新"
python <resources_path>/scripts/minimal_revision.py --active --replace "旧" "新"
```

## 工作流程

```
- [ ] 确认 Word + pywin32
- [ ] 写稿 → COM 直出会话目录 .docx（author）；勿默认 md→docx
- [ ] 仅当用户已有 MD / 明确转换 → md_to_docx.py
- [ ] GetActiveObject 复用；禁止同路径二次 Open
- [ ] DisplayAlerts=0；Visible=True
- [ ] 填表 → parse_profile（若需）→ inspect → plan → auto_fill（--draft，勿默认 --save）→ 用户审阅确认后再定稿
- [ ] 批改 → minimal_revision（默认 Save）；填表默认不定稿；均禁止 AcceptAll
- [ ] 不用 Selection 改内容
```

## COM 硬规则

1. COM 集合 **1-based**；段末 `\r`，单元格 `\x07`。
2. **改已有正文**（含删字）→ 只能 `minimal_revision.py`。**禁止**静默删除。**新建/学校 fill-section** → `MoveEnd(1,-1)` 再 `Text = …`。
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

**路径选型**：见上文 [python-docx 与 Word COM 职责](#python-docx-与-word-com-职责必守)。COM 探查/会话用本节命令；`md_to_docx.py` 不必开 Word。
