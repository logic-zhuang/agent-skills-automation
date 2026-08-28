---
name: Stata命令行分析
description: 仅通过 Stata 命令窗口运行命令、读取结果文本和表格，并通过导出图形及窗口截图获取图形。用户要求用 Stata 分析数据、编写 do 文件、描述统计、回归或导出图形，且不使用菜单、对话框、ui_map 或 roadmaps 时使用。
source: repo
---

# Stata 命令行操作

仅通过命令窗口驱动正在运行的 Stata 会话，读取结果窗格并导出或截图图形。
**不得**使用菜单、对话框、`ui_map` 或 `roadmaps`。

`resources_path` 来自 `load_skill("Stata命令行分析")`（本目录；技能目录名仍为 `stata-cli`）。脚本路径用 `/`。说明见 [README.md](README.md)。

## 稳定反馈契约（摘要）

全体系见 [SKILLS_SYSTEM.md](../SKILLS_SYSTEM.md)。分析时：

1. **入口**：`/stata`、回归/描述统计/do → 本 skill；勿与 Soft Pack `stata_ui` 混用。
2. **禁跳步**：新任务先 `new-task`；一次一条 `run`；等 JSON 返回再下一条；用 `materials search` 查语法，不整文件灌进上下文。
3. **回复骨架**：判定任务 → 已跑命令摘要 → 关键数字 + `out_dir` 图路径 → 待确认假设/下一步模型；**不要**整窗 Results 粘贴。

## 快速流程

```text
python <resources_path>/scripts/cli.py new-task --workspace <abs session dir>
python <resources_path>/scripts/cli.py run "sysuse auto, clear"
python <resources_path>/scripts/cli.py run "summarize mpg weight"
python <resources_path>/scripts/cli.py run "scatter mpg weight"
python <resources_path>/scripts/cli.py read-results
python <resources_path>/scripts/cli.py materials search regress
```

**Every new user task** starts with `new-task` (or `connect --clear` / `clear-panes`):
`clear all` (variables + properties) → `cls` (Results) → clear History → optional `graph drop _all`.

Each call prints JSON (`ok`, `output`, `tables`, `graphs`, `out_dir`). Default `out_dir` is `<cwd>/.coworker/stata_cli`. Graphs land in `out_dir/graphs/`.

If Stata is not running, `connect` / `new-task` / `run` launch it (`STATA_EXE` or `--exe`, else common install paths) and wait until the main window is ready.

## 硬性规则

- **Never run two `cli.py connect` / `run` / `read-*` at once** (parallel typing corrupts commands). Wait for each JSON result before the next. A session file lock enforces this.
- One Stata command per `cli.py run` (or one absolute `do "C:/abs/file.do"`).
- Absolute paths + `/` for `do` / `save` / `import` / `export` / `graph export`. Relative names often hit `r(601)`.
- Command source = [materials.yaml](materials.yaml). **Never** dump the whole file into the prompt.
- Do **not** click menus, fill dialogs, or call Soft Pack `stata_ui_demo` / `coworker.ui_executor`.
- If Soft Pack `stata` tools (`stata_run`, …) are also enabled, pick **one** stack for the session — do not mix.
- Reply with short numbers + graph paths; do not paste the entire Results window.

## 渐进式查阅（命令字典）

1. Index (family counts + ids): `python …/scripts/cli.py materials index`
2. Syntax / example: `python …/scripts/cli.py materials search <query>` (optional `--family stats --limit 4`)
3. Ambiguous / missing option → in Stata: `help <cmd>`
4. Cwd / do-file / graph notes → [reference.md](reference.md)

Search examples: `merge`, `regress`, `summarize`, `graph export`, `ttest`.

## 工作流程

1. **`new-task`** (required at the start of each new user request) — clears Results / Variables / Properties / History before any analysis.
2. Load data (`sysuse`, `use`, `import delimited` / `import excel`) with absolute paths.
3. Inspect (`describe`, `summarize`, `tabulate`) then transform / model from materials hits.
4. For figures: `scatter` / `twoway` / `histogram` … — runtime auto `graph export` PNG under `out_dir/graphs/` and may screenshot the Graph window.
5. Multi-step work: write a `.do` under the session folder, then `run 'do "C:/abs/analysis.do"'`.
6. Stop after user-facing numbers/files are in the JSON.

Mid-task pane reset (same session, same user task still going): `clear-panes` only.

## 环境要求

- Windows + live Stata GUI (MP/SE/BE). Chinese or English UI labels OK.
- Needs `uiautomation` + `pywin32` (see [requirements.txt](requirements.txt)).
- Emergency stop while typing: **Ctrl+Shift+Q** (this process only).
- Optional: `STATA_EXE`, `STATA_CLI_OUT`.
