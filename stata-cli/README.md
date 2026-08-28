# stata-cli

OpenWorker / Cursor agent skill that drives a **live Stata GUI** through the
**Command window** only: run commands, read Results, export/capture graphs.

No menus, no dialogs, no `ui_map`, no roadmaps.

## Requirements

| Item | Notes |
|------|--------|
| OS | **Windows** only |
| Stata | MP / SE / BE with GUI (tested on Stata 19) |
| Python | 3.10+ recommended |
| Packages | `pip install -r requirements.txt` (`uiautomation`, `pywin32`) |
| UI language | Chinese or English pane titles are both supported |

## Install for someone else

### A. Drop into an OpenWorker workspace (recommended)

1. Copy this whole folder to either:
   - `<workspace>/skills/stata-cli/` (repo-shared; needs OpenWorker that scans `skills/`), or
   - `<workspace>/.coworker/skills/stata-cli/` (always discovered)
2. Install deps into the same Python the agent uses:
   ```text
   pip install -r skills/stata-cli/requirements.txt
   ```
3. Point at Stata if it is not under `C:\Program Files\Stata19|18\...`:
   ```text
   set STATA_EXE=D:\path\to\StataMP-64.exe
   ```
4. Enable the skill in Settings ▸ Skills (name: `stata-cli`), then ask in chat, e.g.
   「用 auto 数据 summarize mpg，再 regress mpg weight」.

You do **not** need to run `cli.py` by hand — the agent loads `SKILL.md` and
runs the scripts via its shell tool.

### C. Agent session test (recommended after smoke)

See [TEST-AGENT.md](TEST-AGENT.md) for copy-paste prompts (L0–L4), acceptance
checks, and failure triage when testing through Cursor / OpenWorker Agent.

### B. Standalone smoke test (optional)

```text
python scripts/cli.py materials search regress --limit 2
python scripts/cli.py connect --exe "%STATA_EXE%" --workspace C:/abs/session
python scripts/cli.py run "sysuse auto, clear"
python scripts/cli.py run "summarize mpg"
```

## Soft Pack `stata` vs this skill

| | Soft Pack `packs/stata` | Skill `stata-cli` |
|--|------------------------|-------------------|
| Tools | `stata_connect` / `stata_run` / `stata_ui_demo` | `scripts/cli.py` via shell |
| Menus / roadmaps | Optional UI demos | **Never** |
| Enable | Soft Pack + pack skill | This folder only |

Prefer **one** path per session. If Soft Pack Stata is already on, either use
that pack or mute it and use `stata-cli` — do not mix both against the same window.

## Agent rules (also in SKILL.md)

- **Start every new user task with `new-task`** (clears Results / variables / properties / History).
- **One `cli.py run` at a time** — a file lock blocks parallel typing into Command.
- Absolute paths with `/` for `do` / `save` / `import` / `graph export`.
- Command syntax from `materials.yaml` via `cli.py materials search` — never dump the whole YAML.
- Graphs land under `out_dir/graphs/` (default `<cwd>/.coworker/stata_cli`).

## Clear panes (new task)

```text
python scripts/cli.py new-task --workspace C:/abs/session
# equivalent:
python scripts/cli.py connect --clear
python scripts/cli.py clear-panes
```

What it does:

| Pane | How |
|------|-----|
| Variables + Properties | `clear all` (drops data in memory) |
| Results | `cls` |
| History | UIA clear (context menu or Delete) |
| Graphs | `capture graph drop _all` |

Destructive for unsaved data in memory — intended when starting a fresh task.

## Environment variables

| Variable | Purpose |
|----------|---------|
| `STATA_EXE` | Full path to `StataMP-64.exe` / SE / BE |
| `STATA_CLI_OUT` | Default session output directory |

## Emergency stop

While a CLI process is typing into Stata: **Ctrl+Shift+Q** exits that process only
(does not kill your shell).

## What not to commit

Session dumps under `.coworker/stata_cli*` may include Results text (license banner,
paths). Keep them local.
