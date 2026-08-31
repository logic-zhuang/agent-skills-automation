# 场景：Excel 填表（Agent 主导）

用户上传**原样** `.xlsx` + 档案。  
系统：`inspect` 报**事实** → **Agent 定策略** → 脚本执行 `write_cell` / `append_row`。

## 分工

| 层 | 负责 |
|----|------|
| `inspect_sheet` | `facts`（含 `layout_hint`）+ 可选 `suggestions` |
| **Agent** | `label_left_value_right` → 填邻格；`table_rows` → `append_row` |
| `fill_sheet` / `auto_fill_sheet --fill` | clone→draft、执行 ops |

左→右启发式**不是**唯一能力；行式成绩表必须由 Agent 用 `append_row`。

## 输出落点

| 环境 | 填写结果 |
|------|----------|
| **OpenWorker / 用户会话** | 会话工作区 |

```text
禁止：Open(原文件) → 改 → SaveCopyAs
正确：复制 → 会话目录/draft.xlsx → 只改 draft
```

## Agent 硬流程

```text
1. parse_profile → profile.json
2. inspect_sheet → facts.layout_hint
3. Agent 写 fill.json（cells 和/或 ops）
4. auto_fill_sheet source.xlsx --fill fill.json --draft <会话>/draft.xlsx
5. 用户核对 draft
```

## fill.json 原语

```json
{
  "schema": "excel-com/fill/v1.1",
  "cells": [{"sheet": "Sheet1", "address": "B2", "text": "张三"}],
  "ops": [
    {
      "op": "append_row",
      "sheet": "Sheet1",
      "values": {"B": "学号", "F": "课程名", "K": "90"}
    }
  ]
}
```

## 命令

```bash
python <resources_path>/scripts/inspect_sheet.py "form.xlsx" --json inspect.json
python <resources_path>/scripts/auto_fill_sheet.py "form.xlsx" --fill fill.json --draft draft.xlsx --highlight

# 旧登记表启发式
python <resources_path>/scripts/auto_fill_sheet.py "form.xlsx" --kb profile.json --auto-suggest --draft draft.xlsx
```
