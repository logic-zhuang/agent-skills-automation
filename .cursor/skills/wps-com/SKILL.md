---
name: wps-com
description: >-
  Windows WPS 文字 COM 执行层：打开/复用文档、Markdown→docx（md_to_docx）、inspect、
  最小修订（TrackRevisions+diff）、填表与混合文档。不含学校论文格式。用户提到 WPS、
  金山、Kwps、WPS 文字、Markdown 转 Word、填表、润色、批改、改稿、修订、track changes 时使用。
source: repo
---

# WPS COM（执行层）

Windows + WPS Office 文字 + `pywin32`。**只负责怎么动 WPS**，不含学校/院系格式规则。

与 [word-com](../word-com/SKILL.md) **产品能力对齐**（MD→docx / 润色 / 填表 / hybrid），但 **仅挂接 WPS**（`Kwps.Application` / `wps.Application`），**不回退 Microsoft Word**。

论文套版默认仍可走 `word-com` + `*-thesis-format`。若用户明确用 WPS 套版，须 **同时加载本 skill** 与对应学校 skill，并将 COM 脚本路径改为 `wps-com`。

## 与学校 thesis-format skill 配合（可选）

宿主 App / Agent 会话在 WPS 环境下应 **同时加载**：

```text
["wps-com", "cupk-thesis-format", ...]
```

| 任务 | 读哪个 skill | 跑哪个脚本 |
|------|--------------|------------|
| **Markdown → docx**（OpenWorker 写稿） | 本 skill → [scenarios/md-to-docx.md](scenarios/md-to-docx.md) | `scripts/md_to_docx.py` |
| 润色 / 批改 / 改字改段 | 本 skill → [scenarios/revise.md](scenarios/revise.md) | `scripts/minimal_revision.py` |
| 填固定表格 | 本 skill → [scenarios/form-fill.md](scenarios/form-fill.md) | `inspect` → Agent `fill.json` → `auto_fill_form --fill --draft` |
| 表 + 分段框架（申请书等） | 本 skill → [scenarios/hybrid.md](scenarios/hybrid.md) | `scripts/hybrid_ops.py` |
| 某校论文套版 | **学校 skill**（COM 路径改 wps-com） | 学校 `scripts/convert_ops.py` 等 |

OpenWorker：`load_skill("wps-com")` 得 `resources_path`；**用户来稿只读**；填写/修订输出写在**会话工作区**。

**访问规则（必守）**：shell cwd 保持会话工作区；用 `"<resources_path>/scripts/..."` 跑脚本。`load_skill` 后可用 `list_files` / `read_file` 读技能目录；若仍报「超出访问范围」，改用 shell，**不要** `cd` 到 `resources_path`。

## 填表分工（Agent 主导）

| 层 | 负责 |
|----|------|
| `inspect_form` | `facts` + 可选 `suggestions` |
| Agent | 判版式、写/改 `fill.json`；`--auto-suggest` 仅旧表回归 |
| `auto_fill_form --fill` / `fill_form` | clone→draft、按计划写入 |

详见 [scenarios/form-fill.md](scenarios/form-fill.md)。

## 场景判定（Agent 先看用户意图）

按顺序匹配，命中即停：

1. **Markdown 转 docx / 只有 .md 要出 docx** → [scenarios/md-to-docx.md](scenarios/md-to-docx.md)
2. **改已有正文**（润色、批改、改几个字、改第 N 段）→ [scenarios/revise.md](scenarios/revise.md)
3. **只填表格字段**（导员表、几项 cover 式字段）→ [scenarios/form-fill.md](scenarios/form-fill.md)
4. **论文套版 / 某校格式 / 毕设排版**→ **转已加载的 `*-thesis-format` skill**（WPS 环境须同时加载本 skill；来稿若是 MD，先 `md_to_docx`）
5. **上半表格 + 下面分部分正文**（申请书）→ [scenarios/hybrid.md](scenarios/hybrid.md)
6. **不确定** → `scripts/inspect_doc.py` 探查后再判

易混：

- 「把 md 转成 word」→ **md-to-docx**，不是套版
- 「改第 2 章表述」→ **批改**（revise），不是套版
- 「套成石大本科格式」→ 若只有 MD：先 md-to-docx，再 **学校 skill**（cupk），COM 用 wps-com
- 用户说 Word 但未装 Word → 提示用本 skill（WPS）或安装 Word 后用 word-com

## 必读

| 文档 / 脚本 | 用途 |
|-------------|------|
| [reference.md](reference.md) | WPS ProgID、COM 对象、与 Word 差异 |
| [minimal-revision.md](minimal-revision.md) | 最小修订策略详解 |
| [scenarios/md-to-docx.md](scenarios/md-to-docx.md) | OpenWorker MD → 粗 docx |
| [scenarios/revise.md](scenarios/revise.md) | 批改流程 |
| [scenarios/form-fill.md](scenarios/form-fill.md) | 填表：结构报告 → plan → 修订草稿 → 用户确认再定稿 |
| [kb-profile.md](kb-profile.md) | 知识库 L1/L2；用户不写 JSON |
| [scenarios/hybrid.md](scenarios/hybrid.md) | 混合文档流程 |
| [scripts/md_to_docx.py](scripts/md_to_docx.py) | Markdown → .docx（python-docx，不必开 WPS） |
| [scripts/wps_session.py](scripts/wps_session.py) | attach_wps / open_doc / clone_for_edit |
| [scripts/auto_fill_form.py](scripts/auto_fill_form.py) | 填表流水线（默认不定稿 + TrackRevisions） |
| [scripts/inspect_form.py](scripts/inspect_form.py) | 结构报告 v2 |
| [scripts/plan_fill.py](scripts/plan_fill.py) | slots + profile → fill；needs_confirm |
| [scripts/parse_profile.py](scripts/parse_profile.py) | 自然语言/键值 → 内部 profile |
| [scripts/fill_form.py](scripts/fill_form.py) | COM 写入 cells |
| [scripts/form_utils.py](scripts/form_utils.py) | 结构探测 / 别名 / verify |
| [scripts/minimal_revision.py](scripts/minimal_revision.py) | 最小 diff 修订写回 |
| [scripts/inspect_doc.py](scripts/inspect_doc.py) | 只读探查 |
| [scripts/com_session.py](scripts/com_session.py) | 打开并保持 WPS 窗口 |

## 默认改稿策略（改已有正文）

不必等用户说「最小修订」：

- 用 `minimal_revision.py` + `TrackRevisions=True` + `SequenceMatcher` diff
- 默认 **Save**（`--no-save` 才跳过）
- **不**自动 `AcceptAll`
- **禁止**整段 `Paragraph.Range.Text = 新全文` / `doc.Content.Text = 新全文`

```bash
python <resources_path>/scripts/minimal_revision.py "doc.docx" --paragraph 2 --new "润色后段落"
python <resources_path>/scripts/minimal_revision.py "doc.docx" --replace "旧" "新"
python <resources_path>/scripts/minimal_revision.py --active --replace "旧" "新"
```

演示：`python wps-com/examples/md-to-docx/run_demo.py`（MD→docx）  
演示：`python wps-com/examples/revise/demo_minimal_revision.py`（润色）  
演示：`python wps-com/examples/form-fill/scholarship-cnpc/run_scholarship_demo.py`（填表）  
用例索引：[examples/README.md](examples/README.md)

## 工作流程

```
- [ ] 确认 WPS Office + pywin32（未装 WPS 则报错，不切 Word）；MD 转 docx 另需 python-docx
- [ ] 仅有 Markdown → md_to_docx.py → 会话目录 .docx，再进入后续场景
- [ ] attach_wps：GetActiveObject 复用；禁止同路径二次 Open
- [ ] DisplayAlerts=0；Visible=True
- [ ] 填表 → parse_profile（若需）→ inspect → plan → auto_fill（--draft，勿默认 --save）→ 用户审阅确认后再定稿
- [ ] 批改 → minimal_revision（默认 Save）；填表默认不定稿；均禁止 AcceptAll
- [ ] 不用 Selection 改内容
```

## COM 硬规则

1. COM 集合 **1-based**；段末 `\r`，单元格 `\x07`。
2. **改已有正文** → `minimal_revision.py`。**新建/fill-section** → `MoveEnd(1,-1)` 再 `Text = …`。
3. 中英混排：`Font.NameFarEast` + `Font.NameAscii`；不设 `NameOther`。
4. 样式名带空格：`"标题 1"`、`"正文"`（WPS 中文版与 Word 通常一致，以 `inspect_doc` 为准）。
5. 页码：`Footers(i).PageNumbers.RestartNumberingAtSection`；不写 `PageSetup.RestartPageNumbering`。
6. 奇偶页眉：先 `LinkToPrevious = False`。
7. 分节：`InsertBreak(2)`，不要只用分页符。
8. OMath：WPS 上可能不可用或行为不同 → 见 [reference.md](reference.md) 限制节。
9. 跳过 `~$*.docx`；同文件勿二次 Open。

## 探查

```bash
python <resources_path>/scripts/inspect_doc.py "path/to.docx"
python <resources_path>/scripts/com_session.py "path/to.docx"
```

stdout 应含 `backend: Kwps.Application` 或 `wps.Application`。

## 与 word-com 选型 / 能力对齐

| 用户环境 | 加载 skill |
|----------|------------|
| 仅 Microsoft Word | `word-com` |
| 仅 WPS | `wps-com` |
| 两者都有且未指定 | 按用户说法；本 skill 永不绑 Word |

| 能力 | word-com | wps-com |
|------|----------|---------|
| MD → docx | `md_to_docx.py` | **同脚本** |
| 润色 / 填表 / hybrid | 有 | **有**（`wps_session`） |
| 学校套版 COM | 默认 | 可选（改路径） |
| Word 格式演示脚本 | `demo.md` 等 | 不复制（非产品主路径） |

页眉奇偶、分节、域、OMath、TrackRevisions、可见窗口 → 必须 COM（WPS 或 Word 二选一）。  
**Markdown → 粗 docx**：`scripts/md_to_docx.py`（不必开 WPS）。
