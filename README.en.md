# desktop-rip

Windows desktop automation **Agent Skills**: drive Word, WPS Writer, Excel, and Stata via COM/CLI, plus university thesis formatting packs.  
Built for OpenWorker / Cursor agents—**how to operate desktop apps**, not school-specific layout rules (those live in thesis-format skills).

[中文](README.md) | English

## Skills

| Directory | Purpose |
|-----------|---------|
| [word-com](word-com/SKILL.md) | Word COM: MD→docx, revise, form-fill, hybrid docs |
| [wps-com](wps-com/SKILL.md) | WPS Writer COM (display name: WPS文档操作; parity with word-com; never falls back to Word) |
| [excel-com](excel-com/SKILL.md) | Excel COM: inspect → draft fill |
| [stata-cli](stata-cli/SKILL.md) | Stata via command window only |
| [cupk-thesis-format](cupk-thesis-format/SKILL.md) | CUPK undergrad thesis layout (requires word-com) |
| [thesis-format-skill-builder](thesis-format-skill-builder/SKILL.md) | Author new school thesis-format skills |

System contract (entry prefixes, pipelines, reply shape, stability): **[SKILLS_SYSTEM.md](SKILLS_SYSTEM.md)** (Chinese).

## Architecture

```text
User intent (/word /wps /excel /stata /thesis or natural language)
        │
        ▼
   SKILL.md scenario router (first match wins)
        │
        ▼
   scripts/*.py  (pywin32 COM / Stata CLI)
        │
        ▼
   Session workspace artifacts  ← user source files stay read-only
```

## Requirements

- Windows 10/11  
- Python 3.10+, `pywin32`  
- Optional apps: Microsoft Word / WPS Writer / Excel / Stata  
- MD→docx: `python-docx`

```bash
pip install pywin32 python-docx
```

## OpenWorker install

Published bundle:

https://gitee.com/cupk-gold-rock_1/skill

```text
git clone https://gitee.com/cupk-gold-rock_1/skill.git
cd skill
python scripts/install_openworker_skills.py --desktop
```

This repo is the **desktop-rip source**; the skill repo stores packs under `sources/desktop-rip/<skill>/`.  
Dev remote: https://gitee.com/lu-family-treasure/desk

## Smoke tests

Use a real user file in the session workspace (inspect → draft → review).  
This package ships **without** `examples/` demos.

## Contributing

1. Keep changes aligned with [SKILLS_SYSTEM.md](SKILLS_SYSTEM.md)  
2. Validate draft flow with a real sample file  
3. Sync `.cursor/skills/` when developing in Cursor  
4. Publish to `cupk-gold-rock_1/skill` `sources/desktop-rip/` when releasing for OpenWorker (no examples/)
