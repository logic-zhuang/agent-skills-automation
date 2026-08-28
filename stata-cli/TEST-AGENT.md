# 在 Agent 里测试 `stata-cli`

本文是给 **Cursor Agent / OpenWorker Agent** 用的测试方案。目标不是“人手工跑 cli.py”，而是验证 **Agent 能否按 SKILL.md 正确加载 skill、串行调用脚本、返回可核对的结果**。

工作区示例：`D:/Desktop/desktop-rip`

---

## 1. 一次性前置条件

| 项 | 要求 | 自检 |
|----|------|------|
| OS | Windows | — |
| Stata | MP/SE/BE，带 GUI | 能手动打开 Stata |
| Python | 3.10+，Agent 用的同一个解释器 | `python --version` |
| 依赖 | `uiautomation`, `pywin32` | `pip install -r .cursor/skills/stata-cli/requirements.txt` |
| Skill 路径 | `.cursor/skills/stata-cli/` 存在 | 已复制到 Cursor 技能目录 |
| Skill 开关 | Cursor Settings → Skills → 启用 `stata-cli` | 新开会话后再测 |
| Stata 路径 | 非默认安装时设环境变量 | `$env:STATA_EXE = "C:/Program Files/Stata19/StataMP-64.exe"` |
| 工作区 | Agent 会话 workspace = 仓库根 | `D:/Desktop/desktop-rip` |

**注意**

- 测试期间 **Stata 窗口不要最小化到无法 UIA 聚焦**；可以后台，但别关。
- 不要同时启用 Soft Pack `stata` 与 `stata-cli` 操控同一 Stata 窗口。
- `.coworker/stata_cli/` 是 Agent 会话产物，可本地保留，勿提交 git。

---

## 2. Agent 测试原则

Agent 应遵守 SKILL.md 的硬规则：

1. **每个新用户任务** 先 `new-task --workspace <abs>`
2. **一次只跑一条** `cli.py run`（或一条 `do "C:/abs/file.do"`）
3. **禁止** 点菜单、填对话框、并行多个 `run`
4. 命令语法先 `materials search`，不全文 dump `materials.yaml`
5. 回复给用户：**关键数字 + 图路径**，不粘贴整屏 Results

你要观察的不只是“Stata 算对了”，还有 **Agent 是否按 skill 流程操作**。

---

## 3. 推荐测试顺序

```text
L0  环境探针（不碰 Stata）
L1  连接 + 清窗 + 基础分析（smoke）
L2  多步分析 + 图导出 + materials 查词
L3  do-file 批量任务
L4  负向 / 边界（验证 Agent 不犯规）
```

每个级别 **开新 Agent 会话** 测一次，避免上一条任务的 Stata 状态干扰。

---

## 4. 测试用例与提示词

### L0 — 环境探针（期望：不启动 Stata）

**提示词（复制到 Agent 聊天）：**

```text
请使用 stata-cli skill，只做环境检查，不要连接 Stata：
1. 列出 materials index 的 family 数量
2. materials search regress --limit 2
3. 确认 .cursor/skills/stata-cli/scripts/cli.py 存在
把 JSON 里的 ok 和搜到的命令 id 告诉我。
```

**验收**

- [ ] Agent 调用 `python .cursor/skills/stata-cli/scripts/cli.py materials index`
- [ ] Agent 调用 `materials search regress --limit 2`
- [ ] **没有** 调用 `new-task` / `connect` / `run`
- [ ] 返回里有 `regress` 相关条目

---

### L1 — Smoke：加载数据 + 描述统计

**提示词：**

```text
使用 stata-cli，在 D:/Desktop/desktop-rip/test/stata-session 开始一个新任务：
1. new-task
2. 加载 auto 数据
3. summarize mpg weight
请只通过 Command 窗口操作，不要点菜单。
返回 mpg 和 weight 的 Obs、Mean，并说明 out_dir 路径。
```

**验收**

- [ ] 第一条命令是 `new-task --workspace D:/Desktop/desktop-rip/test/stata-session`（或等价绝对路径）
- [ ] `run "sysuse auto, clear"` 成功，`ok: true`
- [ ] `run "summarize mpg weight"` 返回表格，`mpg` Mean ≈ 21.30，`weight` Mean ≈ 3019.46
- [ ] JSON 含 `out_dir`，默认 `.../.coworker/stata_cli`
- [ ] `.coworker/stata_cli/session.jsonl` 新增记录

**人工核对（可选）**

```powershell
Get-Content .coworker\stata_cli\session.jsonl -Tail 3
```

---

### L2 — 回归 + 散点图 + 查 materials

**提示词：**

```text
继续用 stata-cli（如果这是新会话，请先 new-task，workspace 同上）：
1. sysuse auto, clear
2. regress mpg weight
3. scatter mpg weight
4. 用 materials search 查 graph export 的用法（limit 2）
汇报：weight 系数、R-squared、散点图 PNG 的完整路径。
```

**验收**

- [ ] `regress` 返回 `weight` 系数约 -0.006，`R-squared` 约 0.65
- [ ] `scatter` 后 JSON `graphs` 非空
- [ ] 存在 `.coworker/stata_cli/graphs/NNN_export.png`
- [ ] Agent 调用了 `materials search graph export`
- [ ] 最终回复简短，**没有** 粘贴整段 Results

---

### L3 — do-file 多步任务

**前置：** 在 `test/stata-session/analysis.do` 放如下内容（可让人工或 Agent 先创建）：

```stata
set more off
sysuse auto, clear
regress mpg weight foreign
predict r, residuals
scatter r weight, yline(0) title("Residuals vs weight")
```

**提示词：**

```text
使用 stata-cli 执行 D:/Desktop/desktop-rip/test/stata-session/analysis.do。
要求：
- 先 new-task（新会话时）
- 只用一条 run 调用 do 文件（绝对路径）
- 返回 foreign 系数是否显著、残差图路径
```

**验收**

- [ ] 仅 **一条** `run 'do "D:/Desktop/desktop-rip/test/stata-session/analysis.do"'`
- [ ] 没有拆成多条 `regress` / `predict` / `scatter` 分别 run（除非 do 不存在）
- [ ] 图导出成功
- [ ] Agent 能解读 `foreign` 的 p 值（约 0.13，不显著）

---

### L4 — 负向测试（验证 Agent 守规矩）

#### L4a 禁止并行

**提示词：**

```text
用 stata-cli 同时跑两条命令：summarize mpg 和 summarize weight。
如果 skill 不允许并行，请按规则串行执行并说明原因。
```

**验收**

- [ ] Agent **串行** 两次 `run`，或明确拒绝并行
- [ ] 没有出现两个同时进行的 `cli.py run` 进程

#### L4b 禁止菜单

**提示词：**

```text
用 stata-cli 通过 Stata 菜单 Statistics → Summarize 来查看 mpg。
如果 skill 不允许，请改用 Command 窗口完成同等分析。
```

**验收**

- [ ] Agent 拒绝点菜单，改用 `summarize mpg`
- [ ] 没有调用 `stata_ui_demo` / UI 自动化菜单方案

#### L4c 新任务必须清窗

**提示词（连续两轮）：**

```text
第一轮：new-task 后 sysuse auto, clear，然后 generate x = 1。
第二轮（同一聊天里新开一个“用户任务”语气）：直接 summarize x，不要 new-task。
```

然后你再发：

```text
这是新任务：请 summarize mpg。先说明会不会 new-task，再执行。
```

**验收**

- [ ] 第二个独立任务前，Agent 应 **先 `new-task`**
- [ ] 若未 new-task，`summarize x` 可能因 `clear all` 失败——Agent 应能发现并纠正

---

## 5. 通过标准（Agent 层）

| 级别 | 通过条件 |
|------|----------|
| L0 | materials 命令可用，未误连 Stata |
| L1 | new-task → sysuse → summarize 全绿，数字正确 |
| L2 | 回归 + 图 + materials search，路径与数字正确 |
| L3 | 单条 do 跑通多步，图导出成功 |
| L4 | 串行、禁菜单、新任务清窗 三条规则均遵守 |

**整体通过：** L0–L3 全过，且 L4 至少 2/3 子项符合预期。

---

## 6. 失败时怎么定位

| 现象 | 先看 |
|------|------|
| `Stata main window not found` | Stata 是否已开；`STATA_EXE` 是否正确 |
| `stata-cli busy` | 是否有并行 `cli.py`；等上一条 JSON 结束 |
| `ok: false` 但 Stata 有输出 | `.coworker/stata_cli/session.jsonl` 最后一条 |
| 图没导出 | 命令是否匹配 graph 启发式；手动 `graph export` 是否可行 |
| Agent 不用 skill | Settings 是否启用；会话 workspace 是否为仓库根 |
| Agent 手敲错误路径 | 是否用了 `.cursor/skills/stata-cli/scripts/cli.py` 而非旧路径 |

**紧急停止：** CLI 正在往 Stata 打字时按 **Ctrl+Shift+Q**（只停当前 CLI 进程）。

---

## 7. 建议记录模板

每轮 Agent 测试记一行：

```text
日期 | 级别 | 提示词摘要 | ok | 问题
2026-08-21 | L2 | auto 回归+散点 | pass | —
```

通过 L1–L3 后，即可认为 **stata-cli 在 Agent 内可用**；L4 用于回归测试 Agent 是否持续遵守 skill 规则。

---

## 8. 与手工 smoke test 的关系

| 方式 | 目的 |
|------|------|
| 手工 `cli.py`（见 README §B） | 验证 Python/Stata/UIA 环境 |
| **本文 Agent 测试** | 验证 Skill 文档 + Agent 行为 + 端到端交付 |

环境不通先跑 README 手工 smoke；环境已通（你已有 `session.jsonl` 记录）可直接从 **L1** 开始测 Agent。
