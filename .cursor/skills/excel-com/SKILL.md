---
name: Excel表格操作
description: Windows Excel COM 执行层：打开/复用工作簿、inspect 工作表结构、按档案对齐并填写单元格（默认草稿审阅后再定稿）。不含业务表单语义。用户提到 Excel 填表、xlsx 登记表、工作簿自动化、Excel COM 时使用。
source: repo
---

# Excel表格操作（执行层）

Windows + Microsoft Excel + `pywin32`。**只负责怎么动 Excel**，不含具体学院/业务表规则。

与 **`Word文档操作` 填表产品流程对齐**：**inspect 报事实 → Agent 定策略 → 脚本执行写入**。  
用户**不写** JSON；`inspect.json` / `fill.json` 仅 Agent/调试产物。

OpenWorker：`load_skill("Excel表格操作")` 得 `resources_path`；**用户来稿只读**；填写结果写在**会话工作区**。技能目录名仍为 `excel-com`。

**访问规则（必守）**：shell cwd 保持会话工作区；脚本用 `"<resources_path>/scripts/..."`。`load_skill` 后可读技能目录；若 `list_files` 失败则改用 shell，勿 `cd` 到 `resources_path`。

## 填表分工（必守）

| 层 | 负责 |
|----|------|
| inspect | `facts` + 可选 `suggestions` |
| Agent | 判 `layout_hint`（左标签右空 / `table_rows`），写 `fill.json` |
| fill | `--draft` 先 clone；`write_cell` / `append_row` |

## 输出硬规则（与 word-com 相同）

```text
用户原文件（只读）
    → 复制到会话工作区 draft.xlsx
    → 只打开并填写 draft
    → Save draft；用户确认后再定稿
禁止：先 Open 原文件再改再另存
```

| 环境 | 原文件 | 填写输出 |
|------|--------|----------|
| OpenWorker / 用户会话 | 用户上传 / 绝对路径 | **会话工作区**（临时对话文件夹） |

## 稳定反馈契约（必守）

体系说明见仓库 [SKILLS_SYSTEM.md](../SKILLS_SYSTEM.md)。本 skill 执行时：

1. **入口**：用户说 `/excel`、xlsx、工作簿填表 → 用本 skill；`.docx` 表 → **转** word-com / wps-com。
2. **场景表命中即停**；缺工作表名 / 路径 → 先问或先 `inspect_sheet`。
3. **禁跳步**：parse_profile → inspect → Agent fill → `--draft` 克隆填写；**禁止**对用户原路径原地写。
4. **回复骨架**（每次）：判定场景 → 已做 → 产物路径（draft.xlsx）→ 待确认 → 下一步。
5. Excel **无** Word 级 TrackRevisions；用 draft + 确认代替。

## 场景判定（Agent 先看用户意图）

按顺序匹配，命中即停：

1. **只填工作表字段**（登记表、信息采集表、固定格）→ [scenarios/form-fill.md](scenarios/form-fill.md)
2. **不确定结构** → `scripts/inspect_sheet.py` 探查后再判
3. **Word 表 / 论文** → 转 `word-com` / `*-thesis-format`，不要用本 skill

易混：

- 「填这份 .xlsx」→ 本 skill
- 「填这份 .docx 里的表」→ `word-com`
- Stata 结果导出 Excel → `stata-cli`（`putexcel`），不是本 skill

## 必读

| 文档 / 脚本 | 用途 |
|-------------|------|
| [reference.md](reference.md) | Excel COM 对象、常量、坑 |
| [scenarios/form-fill.md](scenarios/form-fill.md) | 填表：inspect → plan → 草稿 → 确认再定稿 |
| [kb-profile.md](kb-profile.md) | 档案 L1；用户不写 JSON |
| [scripts/auto_fill_sheet.py](scripts/auto_fill_sheet.py) | 填表流水线（默认不定稿） |
| [scripts/inspect_sheet.py](scripts/inspect_sheet.py) | 结构报告 |
| [scripts/plan_fill.py](scripts/plan_fill.py) | slots + profile → fill；needs_confirm |
| [scripts/parse_profile.py](scripts/parse_profile.py) | 自然语言/键值 → 内部 profile |
| [scripts/fill_sheet.py](scripts/fill_sheet.py) | COM 写入单元格 |
| [scripts/sheet_utils.py](scripts/sheet_utils.py) | 结构探测 / 别名 / verify |
| [scripts/excel_session.py](scripts/excel_session.py) | 打开并复用工作簿 |
| [scripts/com_session.py](scripts/com_session.py) | 打开并保持 Excel 窗口 |

## 工作流程

```
- [ ] 确认 Excel + pywin32
- [ ] GetActiveObject 复用；禁止同路径二次 Open
- [ ] DisplayAlerts=0；Visible=True
- [ ] 填表 → parse_profile → inspect → `auto_fill ... --draft <会话目录/draft.xlsx>`（先复制再填）
- [ ] **禁止**对用户原路径原地写入；`--save` 仅用于已是工作副本的路径
- [ ] 写入前跳过公式格；不用 Selection；用 Range / Cells
```

## COM 硬规则

1. COM 集合与行列 **1-based**（`Worksheets(1)`、`Cells(1,1)`）。
2. 读写值用 `Range.Value2`（或 `Value`）；不要用 `Selection`。
3. **同路径勿二次 Open**：先在 `Workbooks` 里按 `FullName` 查找。
4. 跳过 Excel 锁文件 `~$*.xlsx`。
5. 合并单元格：写入左上角（`MergeArea.Cells(1,1)`）；inspect 要标 `merged`。
6. 有公式的格默认 **不覆盖**（`HasFormula`）。
7. 默认不定稿：`--draft` SaveAs；用户确认后再 `--save`。
8. Excel **无** Word 级 TrackRevisions；草稿审阅用 SaveAs + 可选高亮/变更清单。
9. 每步结束不要 `Quit` Excel（保持窗口给用户看）。

## 探查

```bash
python <resources_path>/scripts/inspect_sheet.py "path/to.xlsx"
python <resources_path>/scripts/com_session.py "path/to.xlsx"
```

## 与 openpyxl

简单读写可用 `openpyxl`；**可见窗口审阅、数据验证、复杂打印区域、与已打开簿复用** → 必须 COM。本 skill 默认 COM。
