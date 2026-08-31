# 桌面 Skills 体系：稳定反馈契约

面向 OpenWorker / Cursor Agent：**用户随便说话时，仍走同一条可预期路径。**  
稳定 = **入口稳 + 流水线稳 + 产物稳 + 校验稳**。

本文件是体系总说明；各 skill 的 `SKILL.md` 内嵌同构契约，执行时以对应 skill 为准。

## 1. 入口约定（前缀 / 意图）

自然语言可带域前缀；**有前缀时不得跨域猜**：

| 用户说法（示例） | 加载 skill | 目录名 |
|------------------|------------|--------|
| `/word …`、Word、docx 排版 | Word文档操作 | `word-com` |
| `/wps …`、金山、WPS 文字 | WPS文档操作 | `wps-com` |
| `/excel …`、xlsx 填表 | Excel表格操作 | `excel-com` |
| `/stata …`、回归、do 文件 | Stata命令行分析 | `stata-cli` |
| `/thesis …`、石大/CUPK 套版 | CUPK本科论文套版 **+** Word文档操作 | `cupk-thesis-format` |

无前缀时：按各 skill「场景判定」表**顺序匹配、命中即停**。  
Word / WPS 都可能时：**先问一句**，不得静默回退到另一套 COM。

## 2. 场景 → 固定流水线（禁跳步）

| 场景 | 唯一主路径 | 关键产物 |
|------|------------|----------|
| MD→docx | `md_to_docx.py` | 会话目录 `*.docx` |
| 润色 | 读 old → LLM new → `expected_old` → `minimal_revision` → Save | 修订痕迹在稿；**不** AcceptAll |
| 填表 | parse_profile → inspect → Agent `fill.json` → `--draft` 填写 → 用户确认 → `--save` | `inspect.json` / `fill.json` / `draft.*` |
| hybrid | hybrid.json → `hybrid_ops.py` | stdout 计数 |
| 论文套版 | inspect_source → clone 模板 → cover / fill-section… | 会话目录输出 docx |
| Stata | `new-task` → `run` → `read-results` | JSON + `out_dir` |

**禁止**：跳过 inspect 直接填；对用户原路径原地改再另存；未转 docx 就宣称「已套版完成」。

## 3. 统一回复骨架（每次给用户）

```text
1. 判定：场景=…；后端=Word|WPS|Excel|Stata
2. 已做：…（脚本名 / 关键结果）
3. 产物：会话工作区相对或绝对路径
4. 待确认：…（不确定字段 / 是否定稿 / 是否 Accept）
5. 下一步：…（一条可执行指令即可）
```

失败时同样五段，第 2 点写清**停在哪、缺什么**，不要假装成功。

## 4. 硬门禁（与代码一致）

| 门禁 | 实现位置 |
|------|----------|
| 原文件只读、draft 克隆 | `*_session.clone_for_edit` / `--draft` |
| 同路径不二次 Open | `find_open_doc` / Workbooks 查找 |
| 润色 expected_old | `minimal_revision.py` |
| 填表默认不定稿 | `auto_fill_*` 无 `--draft`/`--save` 拒写 |
| WPS 不回退 Word | `wps_session.WPS_PROGIDS` |
| 不 AcceptAll | SKILL + revise 场景明文 |

## 5. 包边界（文档必须与代码同构）

| 包 | 负责 | 不负责 |
|----|------|--------|
| `word-com` | Word文档操作：COM + MD→docx + 润色/填表/hybrid | 学校版式条文 |
| `wps-com` | WPS文档操作：同上能力，仅 Kwps/wps | 回退 Word；学校条文 |
| `excel-com` | Excel表格操作：填表流水线 | Word 表、Stata putexcel 语义 |
| `cupk-thesis-format` | CUPK本科论文套版：规则与 convert_ops | 单独开 Word（须同时加载 Word文档操作） |
| `stata-cli` | Stata命令行分析：命令窗驱动 | 菜单/UI map |
| `thesis-format-skill-builder` | 高校论文套版技能制作 | 运行时套版 |

安装清单（skill 仓库）：`scripts/install_openworker_skills.py` 的 `DESKTOP_SKILLS` 须与上表桌面包一致。

## 6. 回归锁稳

改 SKILL / 脚本后，用真实用户文件在会话工作区跑通主路径即可（inspect → draft → 审阅），**不必**依赖 examples 演示目录。  

路径、后端 ProgID、draft 落点与 SKILL 声明不一致 → 视为反馈失稳，先修文档或代码再合并。
