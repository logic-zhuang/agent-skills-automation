# desktop-rip

Windows 桌面自动化 **Agent Skills** 源码仓库：用 COM / CLI 驱动 Word、WPS、Excel、Stata，以及高校论文套版。  
面向 [OpenWorker](https://gitee.com/cupk-gold-rock_1/skill) / Cursor Agent——**只负责怎么动桌面软件**，业务语义由各 skill 场景约定。

中文 | [English](README.en.md)

## 包含的 Skills

| 目录 | 显示名 | 用途 |
|------|--------|------|
| [word-com](word-com/SKILL.md) | Word文档操作 | Word COM：MD→docx、润色、填表、混合文档 |
| [wps-com](wps-com/SKILL.md) | WPS文档操作 | WPS 文字 COM（与 Word文档操作 能力对齐，不回退 Word） |
| [excel-com](excel-com/SKILL.md) | Excel表格操作 | Excel COM：inspect → draft 填表 |
| [stata-cli](stata-cli/SKILL.md) | Stata命令行分析 | 仅通过命令窗口跑 Stata |
| [cupk-thesis-format](cupk-thesis-format/SKILL.md) | CUPK本科论文套版 | 石大克拉玛依校区本科论文格式（依赖 Word文档操作） |
| [thesis-format-skill-builder](thesis-format-skill-builder/SKILL.md) | 高校论文套版技能制作 | 制作新学校套版 skill |

体系约定（入口前缀、流水线、回复骨架、稳定反馈）：**[SKILLS_SYSTEM.md](SKILLS_SYSTEM.md)**

## 架构

```text
用户意图（/word /wps /excel /stata /thesis 或自然语言）
        │
        ▼
   对应 SKILL.md 场景判定（命中即停）
        │
        ▼
   scripts/*.py  （pywin32 COM / Stata CLI）
        │
        ▼
   会话工作区产物（draft / 报告 / 图）← 用户原文件只读
```

- **执行层**（word / wps / excel）：打开复用、inspect、draft 填写、最小修订  
- **学校层**（cupk-thesis-format）：版式规则 + `convert_ops`；COM 遵守 word-com（或用户指定的 wps-com）  
- **稳定反馈**：固定流水线 + draft 门禁 + 统一回复骨架，见 `SKILLS_SYSTEM.md`

## 环境要求

- Windows 10/11  
- Python 3.10+、`pywin32`  
- 按需安装：Microsoft Word / WPS 文字 / Excel / Stata  
- MD→docx：`python-docx`（`word-com` / `wps-com` 的 `md_to_docx`）

```bash
pip install pywin32 python-docx
```

## 在 Cursor 中使用

本仓库根目录或 `.cursor/skills/<name>/` 下的 skill 可被 Agent 自动发现。对话中直接说例如：

- 「用 Word 润色第 2 段」→ `word-com`  
- 「用 WPS 填这份表」→ `wps-com`  
- 「按 CUPK 本科格式套版」→ `cupk-thesis-format` + `word-com`

## 安装到 OpenWorker（skill 发行仓）

正式分发目录在：

https://gitee.com/cupk-gold-rock_1/skill

```text
git clone https://gitee.com/cupk-gold-rock_1/skill.git
cd skill
python scripts/install_openworker_skills.py --desktop
```

本仓库是 **desktop-rip 源码**；发行仓路径为 `sources/desktop-rip/<skill>/`。  
开发仓（本仓库）同步到 Gitee：https://gitee.com/lu-family-treasure/desk

用户安装后只含执行所需文件（`SKILL.md` / `scenarios/` / `scripts/` / `templates/` 等），不含开发用 examples。

## 目录速览

```text
desktop-rip/
├── SKILLS_SYSTEM.md          # 稳定反馈与包边界
├── word-com/                 # Word 执行层
├── wps-com/                  # WPS 执行层
├── excel-com/                # Excel 执行层
├── stata-cli/                # Stata CLI
├── cupk-thesis-format/       # CUPK 套版
├── thesis-format-skill-builder/
└── .cursor/skills/           # Cursor 侧同步副本
```

每个 skill 通常含：`SKILL.md`、`scenarios/`、`scripts/`、可选 `templates/`。

## 参与贡献

1. 改 skill 后对照 [SKILLS_SYSTEM.md](SKILLS_SYSTEM.md) 检查入口与流水线是否仍一致  
2. 用真实样表在会话目录验证 draft 流程  
3. 同步 `.cursor/skills/`（若本地用 Cursor）  
4. 需要发布到 OpenWorker 时，再同步到 `cupk-gold-rock_1/skill` 的 `sources/desktop-rip/`（勿带回 examples）

## License

以各子目录及上游依赖声明为准；发行到 skill 仓时遵循该仓 LICENSE。
