# 场景：填表（产品化流程）

用户上传**原样** Word 表 + 知识库里的**基本信息**。  
系统内部：读结构 → 抽档案 → 对齐 → **带修订写入草稿** → **用户审阅确认后再定稿**。

用户**不写** JSON、不插书签。`slots.json` / `fill.json` 仅 Agent/调试产物。

## 何时用 / 不用

- **用**：奖学金/导员表等「很多人填同一类表」；档案在 OpenWorker KB
- **不用**：论文套版 → 学校 skill；改措辞 → [revise.md](revise.md)；表+多节正文 → [hybrid.md](hybrid.md)

## Agent 硬流程（必守）

```text
1. parse_profile（若档案是自然语言/键值文本）→ profile.json
2. inspect_form → 结构报告 v2.1（fill_mode / targets / coverage）
3. 先看 coverage（与档案无关）：
     unresolved_labels = 像标签但找不到写入位 → 漏扫候选
     unbound_empties   = 空着却没被任何 slot 认领
   禁止用 unmatched_kb（档案多出来的字段）判断扫全
4. plan_fill --print → matched / unmatched_slots / needs_confirm
5. 有 needs_confirm 或关键 unresolved → Agent 追问用户，禁止 silently 瞎填
6. fill / auto_fill：TrackRevisions=True，默认不定稿（勿 --save）
7. 请用户在 Word 审阅中核对；用户明确确认后再 --save 或保存草稿定稿
8. 禁止 AcceptAll；禁止在确认前宣称「填表成功」
```

详见 [kb-profile.md](../kb-profile.md)。

## 命令

```bash
# 档案文本 → 内部 profile（用户不碰 JSON）
python <resources_path>/scripts/parse_profile.py profile.txt --out profile.json --print

# 读结构
python <resources_path>/scripts/inspect_form.py "form.docx" --json slots.json

# 对齐报告
python <resources_path>/scripts/plan_fill.py slots.json --kb profile.json --print

# 写入草稿（默认不定稿 + 修订）
python <resources_path>/scripts/auto_fill_form.py "form.docx" --kb profile.json --plan plan.json --draft draft.docx

# 用户确认后再定稿
python <resources_path>/scripts/auto_fill_form.py "form.docx" --kb profile.json --save
```

一条命令（仍默认不定稿）：

```bash
python <resources_path>/scripts/auto_fill_form.py "form.docx" --kb profile.json --plan plan.json --draft draft.docx
```

## 结构报告 v2.1（inspect）

每个 slot 含：

| 字段 | 含义 |
|------|------|
| `fill_mode` | `single` / `digit_grid` / `year_month_day` / `checkbox` / `inline_*` / `multiline` / `unknown` |
| `targets[]` | 要写入的格坐标（拆格为多格） |
| `confidence` | 0–1 |
| `reason` | 如 `label+right_empty`、`label+below_empty`、`label+right_date_template:ymd` |

邻域优先级（结构驱动，不读档案）：

1. 右邻空 / 占位符 → `single`（或电话/身份证拆格）
2. 右邻「年　月　日」模板 → `year_month_day`
3. 右邻多 □（非是/否）→ `unknown`（`needs_confirm`）
4. 否则看**下邻同列**：空 → `single`；日期模板 → `year_month_day`
5. 仍无写入位 → 记入 `coverage.unresolved_labels`

电话规则：

- 标签像电话 + **1** 空位 → `single`
- 标签像电话 + **7–11** 连续小格 → `digit_grid`
- 身份证类 + **≥15** 小格 → `digit_grid`

`unknown` / 低置信度 → plan 的 `needs_confirm`，**不写入**。

### coverage（扫全指标）

| 字段 | 含义 |
|------|------|
| `unresolved_labels` | 像标签但找不到写入位（漏扫候选） |
| `unbound_empties` | 空/日期模板格未被任何 slot 认领 |
| `label_coverage` | resolved / (resolved+unresolved) |
| `scan_complete_enough` | unresolved 与 unbound 皆为 0 时 true（少见） |

## plan 报告

| 字段 | 含义 |
|------|------|
| `matched` | 已扫到且档案有值 → 可自动填 |
| `unmatched_slots` | 已扫到、档案无值（可空） |
| `needs_confirm` | 结构不清，需确认 |
| `coverage` | 透传 inspect；**用它看漏扫** |
| `unmatched_kb` | 档案有、本表未用；**不等于漏扫**（档案是超集） |
| `review_required` | 恒为 true |

## 审阅与保存

| 行为 | 默认 |
|------|------|
| TrackRevisions | 开启 |
| Save 定稿 | **否**；需 `--save` |
| 草稿 | `--draft out.docx` SaveAs |
| AcceptAll | **禁止** |

## 能力边界

- 常见结构（右邻空、下邻空、年/月/日模板、拆字身份证/电话、是□否□、句中排名）→ 自动填
- 多行家属表、句中续写（「至___」）、复杂分数区 → 仍可能进 `unresolved` / `unbound`
- 陌生/怪表 → 填一部分 + 覆盖报告 + `needs_confirm`
- **不承诺**任意 Word 100% 自动填对
- 不做书签方案；不靠穷举每张表

## 演示

```bash
python word-com/examples/form-fill/scholarship-cnpc/run_scholarship_demo.py
```

更多用例目录约定见 [examples/README.md](../examples/README.md)。