# 场景：Markdown → Word（可选桥，非默认）

**非默认写稿路径。** OpenWorker 写稿默认走 COM 直接写 `.docx`（见 [author.md](author.md)）。  
职责边界见 [SKILL.md](../SKILL.md#python-docx-与-word-com-职责必守)：**python-docx 只负责本场景；转完后必须用 COM 打开/润色/套版。**

用 `python-docx` 做**粗转**，不替代学校排版；转完后用 Word COM 打开/润色/套版。

## 何时用

- 会话或用户目录里已有 `.md`，要变成可打开的 `.docx`
- 用户明确要求 MD→Word，再接润色 / 填表 / `*-thesis-format`

## 何时不用（默认）

- 「写一篇文章 / 生成 Word」→ [author.md](author.md)，**不要**先造 md 再转
- 已有 Word 来稿 → revise / form-fill / cupk

## Agent 硬流程

```text
1. 确认 .md 路径（用户来稿只读时先复制到会话目录）
2. python <resources_path>/scripts/md_to_docx.py article.md --out article.docx --print
3. 需要学校格式 → *-thesis-format，把 docx 当来稿
4. 需要润色 → revise.md
5. 禁止把「写稿」默认做成 md→docx；禁止未转 Word 就宣称已按论文格式排好
```

## 命令

```bash
python <resources_path>/scripts/md_to_docx.py "article.md" --out "article.docx" --print
```

依赖：`pip install python-docx`（若环境未装）。

## 能力边界

| 支持 | 不支持 / 弱支持 |
|------|-----------------|
| `#`～`######` 标题 | HTML、脚注 |
| 段落、空行分段 | 图片 |
| `-` / `1.` 列表（单层） | 嵌套列表 |
| `**粗**` `*斜*` `` `代码` `` | 复杂链接样式 |
| 简单 GFM 表格 | 合并单元格 |
| 围栏代码块 | LaTeX / OMath |

输出是**可编辑草稿**；页眉页码、石大格式由后续 skill 处理。
