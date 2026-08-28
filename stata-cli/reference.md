# Stata CLI notes (cwd, do-file, graphs)

Command syntax lives in `materials.yaml`. This file is usage only. Humans: [README.md](README.md).

## Soft Pack vs this skill

- Soft Pack `stata`: registered tools + optional UI roadmaps.
- This skill: Command window only via `scripts/cli.py`.
- Do not drive the same Stata window with both in one session.

## New task — clear panes first

At the start of each new user task:

```text
python scripts/cli.py new-task --workspace C:/abs/session
```

This runs `clear all` (empty Variables / Properties), `cls` (Results), clears History via UIA, and `graph drop _all`. Same as `connect --clear` or `clear-panes` after an existing connection.

## Working directory

- Stata status bar = `pwd`. Relative `save` / `do` / `graph export` / `import` resolve against it.
- `connect --workspace <dir>` issues `cd` to that folder.
- Prefer absolute paths with `/` anyway (`C:/data/auto.dta`).

```stata
pwd
cd "C:/abs/session"
```

## Do-files

- `do "C:/abs/analysis.do"` echoes commands; `run` does not.
- Relative do-file names often fail with `r(601)`.
- Pass args: `do "C:/abs/file.do" arg1 arg2` (`1', `2' inside the script).
- Put `set more off` near the top of batch scripts.

## Results and tables

- `cli.py run` returns the Results **delta** for that command plus heuristic ASCII `tables`.
- `cli.py read-results` copies the full Results pane when you need a wider tail.

## Graphs

After a graph command the runtime:

1. Runs `graph export "<out_dir>/graphs/NNN_export.png", replace`
2. Screenshots any new Stata Graph window when possible

Deliver those PNG paths. To export yourself:

```stata
graph export "C:/abs/out/fig.png", replace
graph export "C:/abs/out/fig.pdf", as(pdf) replace
```

## Concurrency

Only one `cli.py` process may type into Stata at a time (lock file under `out_dir`).
If you see `stata-cli busy`, wait and retry — never parallelize `run`.

## UI language

Pane titles may be Chinese (`命令窗口` / `结果窗口`) or English (`Command` / `Results`).
Both are resolved automatically. Hotkeys `Ctrl+1`…`Ctrl+5` are used either way.

## What this skill does not do

Menus, dialogs, `ui_map`, `roadmaps`, and UI demos. Stay on the Command window.
