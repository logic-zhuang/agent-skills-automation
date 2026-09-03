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

## 多选 □ 硬规则（必守）

表内常见：`□农应 □农往 □城应 □城往`、`□统一招生 □国家专项 …`

| 正确 | 错误 |
|------|------|
| 只把目标项 `□` 改成 `☑`，**保留全部选项文字** | 删掉其他选项，单元格只剩「城应」 |
| fill 可写短标签 `城应`；`fill_form` 会按当前格模板勾选 | 手写整格覆盖成无 □ 的纯文本 |

引擎：`fill_mode=checkbox_group` + `apply_multi_checkbox`；写入前 `resolve_checkbox_write` 兜底。  
写入时对多选格**暂停 TrackRevisions** 再整格写回（避免修订模式下出现「城应 + 旧选项箭头/重复」）；其它字段仍正常开修订。

## Word 进程硬规则（本 skill）

- **只**通过 `word_session.attach_word` → `Word.Application` 打开/编辑  
- **禁止** `start draft.docx` / 资源管理器打开（会落到系统默认，常为 **WPS**）
- 若用户环境只有 WPS → 转 **wps-com**，不要假装用了 Word
- `attach_word` 会拒绝 Path 落在 Kingsoft/WPS 的伪 `Word.Application`；回复写 `backend: Word.Application` 时对照脚本的 `path=`（须为 Office 的 WINWORD 目录）

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
