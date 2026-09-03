# 场景：Markdown → docx（可选桥，非默认）

**非默认写稿路径。** 写稿默认走 WPS COM（见 [author.md](author.md)）。  
职责边界见 [word-com/SKILL.md](../../word-com/SKILL.md#python-docx-与-word-com-职责必守)：**python-docx 只负责本场景；转完后必须用 WPS COM。**

与 [word-com/scenarios/md-to-docx.md](../../word-com/scenarios/md-to-docx.md) 同一脚本能力；转完后用本 skill 的 WPS COM 打开/润色/填表。

## 何时用

- 已有 `.md`，要变成 `.docx`
- 用户明确「md 转 word / 导出 docx」

## 何时不用（默认）

- 「写一篇 / 生成文档」→ [author.md](author.md)，不要先造 md 再转
- 已有 .doc/.docx → revise / form-fill / cupk

## Agent 硬流程

```text
1. 确认 .md（只读来稿先复制到会话目录）
2. python <resources_path>/scripts/md_to_docx.py article.md --out article.docx --print
3. 套版 → *-thesis-format；COM 用 wps-com
4. 润色 → revise.md
5. 禁止把写稿默认做成 md→docx
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
