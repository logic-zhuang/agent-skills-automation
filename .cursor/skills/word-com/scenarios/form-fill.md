# 场景：填表（Agent 主导）

用户上传**原样** Word 表 + 知识库档案。  
系统：`inspect` 报**事实** → **Agent 定策略 / 出 fill 计划** → 脚本只写入工作副本。

用户**不写** JSON（除非调试）。`slots.json` / `fill.json` 为 Agent/系统产物。

## 分工

| 层 | 负责 |
|----|------|
| `inspect_form` | `facts`（格子/邻域/layout_hint）+ 可选 `suggestions`（旧启发式） |
| **Agent** | 判版式、选写入位、改/写 `fill.json`；不确定则问用户 |
| `fill_form` / `auto_fill_form --fill` | clone→draft、按计划写、TrackRevisions |

`suggestions` / `--auto-suggest` 仅用于已验证旧表回归，**不是**穷举所有新表的主路径。

## 输出落点

| 环境 | 填写结果 |
|------|----------|
| **OpenWorker / 用户会话** | 会话工作区（临时对话文件夹） |

```text
正确：复制原文件 → 会话目录/draft.docx → 只打开 draft 填写
禁止：Open(原文件) → 改 → SaveAs
```

## Agent 硬流程

```text
1. parse_profile → 会话目录/profile.json
2. inspect_form → facts（+ suggestions 可选）+ coverage
3. 先看 facts.layout_hint 与 coverage（勿用 unmatched_kb 判断扫全）
4. Agent 出 fill.json（可采纳/修改 suggestions）；低置信 → 追问
5. auto_fill_form source.docx --fill fill.json --draft <会话>/draft.docx
   （旧 demo：--kb ... --auto-suggest --draft ...）
6. 用户审阅 draft；确认后再对 draft --save；禁止 AcceptAll
```

## 命令

```bash
python <resources_path>/scripts/parse_profile.py profile.txt --out profile.json --print
python <resources_path>/scripts/inspect_form.py "form.docx" --json inspect.json

# Agent 编写 fill.json 后执行
python <resources_path>/scripts/auto_fill_form.py "form.docx" --fill fill.json --draft draft.docx

# 仅建议报告、不写入
python <resources_path>/scripts/auto_fill_form.py "form.docx" --kb profile.json --plan plan.json

# 旧表启发式（回归）
python <resources_path>/scripts/auto_fill_form.py "form.docx" --kb profile.json --auto-suggest --draft draft.docx
```

## inspect v2.2

| 字段 | 含义 |
|------|------|
| `facts` | 每格 text/empty/right_empty/below_empty/…；`layout_hint` |
| `suggestions` / `slots` | 启发式 slot（可选） |
| `coverage` | unresolved_labels / unbound_empties |
