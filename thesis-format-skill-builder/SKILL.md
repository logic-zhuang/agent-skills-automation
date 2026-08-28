---
name: 高校论文套版技能制作
description: 制作和维护高校专用的 Microsoft Word 论文排版技能，包括规范提取、模板资源、转换流程、COM 脚本、验证清单和交付说明。用户要求新做高校论文套版技能、把格式手册做成可复用技能，或扩展 CUPK 等本科论文套版时使用。
---

# 高校论文排版 Skill 制作器

把某所高校、学院或专业的论文格式要求，制作成 **OpenWorker 适配** 的 Agent
Skill。最终产物不是一份排版说明，而是一个能根据来稿执行、检查并交付
`.docx` 的小型工作流。

本仓库 `desktop-rip` 包内的 skill 均按 OpenWorker 约定编写：Agent 通过
`load_skill("<name>")` 取得 `resources_path`，在 **OpenWorker 会话工作区**
运行脚本，技能目录只放规则、模板和脚本，不当作 shell 工作目录。

## OpenWorker 适配（必守）

- `load_skill("<name>")` 返回的 `resources_path` 即本 skill 目录；脚本统一写成
  `python <resources_path>/scripts/...`，路径用 `/`。
- **shell cwd 保持 OpenWorker 默认会话目录**（`OpenWorker/<session-id>/`）；
  不要 `cd` 到 skill 目录。
- 来稿、中间 JSON、输出 `.docx` 写到会话工作区或用户给的绝对路径；模板、
  脚本、规则文件留在 `resources_path` 下读取。
- 同包 skill 用显示名互相引用（如 `Word文档操作`、`CUPK本科论文套版`），不要写
  `.cursor/skills/...` 或仓库根相对路径；文件夹名可仍为英文（如 `word-com`）。
- 需要 Python 依赖时在 skill 目录放 `requirements.txt`，并在 README 说明安装。
- 分发位置：`sources/desktop-rip/<skill-name>/`；用户安装：
  `python scripts/install_openworker_skills.py --desktop`。
- Frontmatter 建议保留 `source: repo`（与 `Stata命令行分析` 等同包 skill 一致）。

## 设计边界

- 每个学校/学位/年份建立独立目录，例如
  `sources/desktop-rip/<school>-thesis-format/`（开发时可放在
  `.cursor/skills/` 或仓库根同名目录）。
- 本 skill 负责“制作排版 skill”，不直接套用某个学校的论文。
- 学校的具体字号、页边距、章节顺序、模板名称等必须来自用户提供的范本、
  格式手册或可核验的官方材料；不能把 CUPK 规则泛化给其他学校。
- Word 文件在 Windows 上优先使用 COM；通用 COM 规则见
  [word-com](../word-com/SKILL.md)。
- 参考 CUPK 的组织方式：主 `SKILL.md` 负责触发和总流程，`rules.md` 负责规范，
  `convert.md` 负责执行细节，`scripts/` 负责脆弱的自动化操作，`templates/`
  保存可直接复制的底板。

## 输入确认

开始制作前，先确认下列信息。缺失的信息标记为“待确认”，不要自行补齐：

- 适用学校、校区、学院、学位层次（本科/硕士/博士）和论文类型；
- 规范版本、发布日期，以及范本/标准模板/填写模板；
- 来稿格式：`.docx`、`.doc`、PDF 转 Word，还是纯文本；
- 是否允许代译英文摘要和图表题注；
- 是否需要复制封面图形、WordArt、域、目录、页眉页脚和公式；
- 目标输出目录、技能目录名和是否需要脚本；
- 本机是否有 Microsoft Word、`pywin32` 和可运行的 Python。

如果用户只说“做一个某校排版 skill”，先收集范本或让用户确认使用哪一版
规范；有足够资料时可以直接先搭骨架，并将未确认项写入缺口清单。

## 交付目录

根据任务需要生成以下结构，未使用的文件不要创建：

```text
<school>-<degree>-thesis-format/
├── SKILL.md
├── rules.md                 # 从权威资料整理出的可执行规则
├── convert.md               # 来稿探查、填充、收口、模板维护与抽查步骤
├── requirements.txt         # 默认含 pywin32；Word COM skill 必备
├── templates/               # 可填写底板（运行时唯一来源；须随 skill 分发）
├── scripts/
│   ├── inspect_source.py    # 只读抽取结构和内容
│   ├── convert_ops.py       # CLI：clone、填充、删节、收口
│   ├── <school>_com.py      # 共享 COM helper（模板路径、页眉、关键词格式等）
│   └── build_standard_template.py   # 可选：从权威范本重建 templates/ 底板
└── examples/                # 版本化的 JSON 夹具（cover.json、ch1.json 骨架）
```

不要把学校规则堆进通用 `word-com` skill；不要为了“完整”复制范本中的教学说明、
打印装订说明或与学生填写无关的页面。

### 模板（必守）

- **可填写模板必须放在 `templates/`**，随 skill 分发；`clone` 只读此目录，**禁止**
  在运行时回退到仓库根、`大学论文模板/` 或其它会话外路径。
- 权威**格式范本**仅用于维护：在 `convert.md` 增加 **§模板维护**，写明范本文件名、
  如何从范本生成可填写底板，以及 `build_standard_template.py` 的用法（有则写命令，
  无则写人工步骤）。范本路径写在维护小节，**不要**写进 `fillable_template_path()`。
- 模板缺失时脚本应 **明确报错**（提示运行维护脚本或手动复制），而不是静默找别处的文件。

### 示例 JSON 放哪（必守）

| 位置 | 放什么 | 用途 |
|------|--------|------|
| `<resources_path>/examples/` | `cover.json`、`ch1.json` 等**最小夹具** | 文档引用、回归测试、Agent 照抄字段结构 |
| OpenWorker **会话工作区** | 来稿、 `extract.json`、输出 `.docx`、用户图片 | 每次套版的运行时产物 |
| 仓库 `example/<skill-name>/`（可选） | 完整端到端演示 | **仅开发用**，不随 OpenWorker 分发 |

`convert.md` 命令示例里的 JSON 路径优先写会话工作区（或绝对路径）；夹具结构引用
`<resources_path>/examples/`。不要把运行时 `extract.json` 或输出 docx 写回 skill 目录。

## 制作流程

### 1. 读取并拆解权威材料

逐页检查范本和格式手册，至少记录：

1. 页面：纸型、单双面、装订、页边距、页眉页脚距边界；
2. 样式：正文、标题层级、目录、摘要、关键词、参考文献的中西文字体、
   字号、颜色、对齐、缩进、行距、段前段后；
3. 结构：封面、中文摘要、英文摘要、目录、正文、结论、文献、附录、致谢
   的节顺序和可选性；
4. 封面：表格/文本框/图片/WordArt 的位置、字段和需删除的提示；
5. 页眉页脚：奇偶页差异、分节继承、页码格式、罗马/阿拉伯起始页；
6. 图、表、公式：编号规则、题注语言、位置、表线、OMath 要求；
7. 目录和域：是否由 Word 自动更新，哪些域必须原样保留；
8. 交付限制：字数、摘要长度、致谢要求、查重或打印要求。

把“范本实际表现”和“批注/文字说明”分开记录。两者冲突时，优先保留证据，
在 `rules.md` 中注明冲突和采用的判定，不要静默猜测。

### 2. 建立可执行规则

`rules.md` 采用稳定的小节顺序：

```text
# <学校> <学位>论文排版规则
## 页面
## 样式
## 结构顺序
## 封面
## 页眉页脚
## 目录
## 摘要与关键词
## 图、表、公式
## 参考文献
## 附录与致谢
## 已知缺口
```

数字必须写成能检查的规则，例如“正文首行缩进 2 字、1.5 倍行距”，而不是
“正文格式规范”。对于每条关键规则说明其来源或观察位置。区分：

- 自动化可保证：样式定义、节、页眉页脚、字段、表格线、图片环绕；
- 只能抽查：内容质量、翻译准确性、单位、引用顺序、分页美观；
- 无法保证：扫描图中的可编辑表格、缺失的官方模板、未安装 Word 的环境。

### 3. 设计转换工作流

`convert.md` 至少包含以下阶段：

```text
1. 探查来稿并提取结构
2. 复制技能自带的可填写模板
3. 填充封面
4. 按节填充摘要、目录和正文
5. 处理章节数与可选节
6. 插入/重建图、表、公式和题注
7. 收口：删提示、更新域、统一样式
8. 抽查并声明交付缺口
```

默认策略是“复制底板后填内容”，不要直接在用户原稿上整体套样式。命令示例一律
用 `<resources_path>`，并说明输入 JSON 的字段：

```bash
python <resources_path>/scripts/inspect_source.py "C:/abs/来稿.docx" --extract extract.json
python <resources_path>/scripts/convert_ops.py clone --out "C:/abs/输出.docx"
```

推荐使用内容块模型，避免脚本依赖原稿段落的偶然编号：

```json
{
  "heading": "第1章  绪论",
  "even_header": "第1章  绪论",
  "blocks": [
    {"style": "标题 2", "text": "1.1  研究目的与意义"},
    {"style": "正文", "text": "正文内容"},
    {"type": "picture", "path": "C:/abs/fig.png",
     "caption_zh": "图1.1  题名", "caption_en": "Fig. 1.1  Title"},
    {"type": "omath", "linear": "x=a+b", "number": "(1.1)"}
  ]
}
```

字段名可以按学校调整，但同一 skill 内必须统一。章节可变时，定义明确的
`delete-section`、`extra-chapter` 或等价操作，并从后往前删除多余节。

### 4. 实现脚本

只为不稳定、重复或容易破坏文档的操作写脚本。脚本至少应：

- 支持绝对路径和带空格的路径；
- 跳过 `~$*.docx` 锁文件；
- 检查输入、输出和已存在文件，必要时显式要求 `--force`；
- 使用 Word COM 时复用已打开文档，禁止同一路径二次 `Open`；
- 修改段落或单元格前保留段落标记/单元格结束符；
- 用 `NameFarEast` 和 `NameAscii` 设置中西文字体；
- 用分节符下一页处理章节，不用普通分页符替代；
- 保留范本页码域，不用 `Fields.Add` 盲目重做页码；
- 对页眉先设置 `LinkToPrevious = False`，再修改奇偶页眉；
- 对公式使用 `OMaths.Add` 和 `BuildUp()`，不要批量删除 `OMaths.Range`；
- 失败时返回明确错误，并尽量不覆盖原文件。

`inspect_source.py` 应该只读，输出至少包括节数、标题、段落、表格、图片、
公式、页眉页脚和封面字段。`convert_ops.py` 的子命令应保持小而可验证，典型
命令为 `clone`、`cover`、`fill-section`、`delete-section`、`extra-chapter`
和 `finish`。若学校模板只需人工操作，不要为了形式生成空脚本。

**共享 COM 模块** `<school>_com.py`（参考 `cupk-thesis-format` 的 `cupk_com.py`）：

- 当 `convert_ops.py` 超过约 300 行，或同一 COM 细节（页眉、关键词、章标题居中、
  模板路径）出现两次以上时，**必须**抽出共享模块；`convert_ops.py` 只做 CLI 编排。
- 共享模块职责：`skill_root()`、`fillable_template_path()`（只读 `templates/`）、
  `attach_word` / `open_or_reuse`、学校特有的 `apply_heading1`、页眉/关键词格式等。
- 文件头 docstring 与 Usage 一律写 `python <resources_path>/scripts/...`，
  **禁止** `.cursor/skills/...` 或仓库根相对路径。
- Word COM skill 默认在 skill 根目录放 `requirements.txt`（至少 `pywin32`）。

### 5. 编写目标 skill 的主文件

目标 skill 的 `SKILL.md` 不超过 500 行，并且只放最常用的决策和命令：

```markdown
---
name: <lowercase-hyphenated-name>
description: <具体说明做什么、何时触发、包含学校/学位/文件类型关键词>
source: repo
---

# <学校>论文套版

<一句话说明适用范围。OpenWorker 通过 load_skill 加载；脚本见下方 resources_path 约定。>

## 路径与 OpenWorker

`resources_path` from `load_skill("<name>")` is this folder. Run scripts with
that path (forward slashes). Shell cwd = OpenWorker session workspace.

| 资源 | 位置 |
|------|------|
| 规则 / 转换细则 | [rules.md](rules.md)、[convert.md](convert.md) |
| 可填写模板 | `templates/...` |
| 脚本 | `<resources_path>/scripts/...` |
| Word COM 基座 | [word-com](../word-com/SKILL.md) |

## 默认策略（必守）
<复制底板、不要改原稿、页码/域/公式等关键禁忌>

## 套版流程
<带 `<resources_path>` 命令的短检查清单>

## 必须 COM
<学校特有的封面、节、页眉页脚、表格、公式操作>

## 交付
<输出文件和必须声明的抽查项>
```

Frontmatter 的 `name` 只能使用小写字母、数字和连字符；`description` 用第三人称，
同时写清“做什么”和“何时使用”。`SKILL.md` 正文必须写明 `resources_path` 与
OpenWorker 会话目录分工，不要把命令写成 `.cursor/skills/...` 或“工作目录为仓库根”。

主文件必须直接链接到 `rules.md`、`convert.md` 和 `word-com`，不要设计多层嵌套
引用。学校名称、版本和适用学位在主文件和规则文件中保持一致。

## 验证清单

制作完成后逐项验证：

- [ ] `SKILL.md` 的 frontmatter 合法，含 `source: repo`，描述包含触发词；
- [ ] 命令示例使用 `<resources_path>`，脚本 docstring 亦无 `.cursor/skills/...`；
- [ ] 明确 shell cwd 为 OpenWorker 会话目录，产物与模板路径分工清楚；
- [ ] `templates/` 已含可填写底板；`fillable_template_path()` 无仓库外回退；
- [ ] `convert.md` 含 §模板维护；`examples/` 仅有 JSON 夹具，运行时产物在会话目录；
- [ ] 存在 `<school>_com.py`（或 `convert_ops.py` 足够短且 COM 逻辑未重复）；
- [ ] `SKILL.md` 少于 500 行，引用文件都存在且只有一层跳转；
- [ ] 规则中的每个关键数字能在范本、手册或用户确认中找到依据；
- [ ] 模板能复制到新输出文件，原稿不会被覆盖；
- [ ] 封面字段、空行、图片和 WordArt 不会误删；
- [ ] 节顺序、奇偶页眉、页码域和起始页码经过检查；
- [ ] 标题、正文、摘要、目录和参考文献样式经过检查；
- [ ] 图、表、公式分别验证，公式为 OMath 而不是普通文本；
- [ ] 脚本对缺字段、缺图片、无 Word、锁文件和重复输出给出可读错误；
- [ ] 用至少一份真实或最小示例来稿跑通 `inspect → clone → fill → finish`；
- [ ] 交付说明列出三线表、公式显示、英文翻译、目录更新等未自动保证项。

## 交付格式

向用户交付时说明：

1. 新建或修改了哪个 skill 目录；
2. 该 skill 适用的学校、学位和规范版本；
3. 模板、规则、转换脚本和验证流程分别在哪里；
4. 已实际验证的步骤；
5. 仍需人工抽查或等待用户提供资料的项目。

不要声称“完全符合”而不列证据。排版 skill 的质量标准是可复用、可追溯、
可验证，而不是规则文字的数量。
