# 场景：Markdown → Word（OpenWorker 桥）

OpenWorker 常见产物是 **Markdown**；后续填表、润色、学校套版都要 **Word**。  
本场景用 `python-docx` 做**粗转**，不替代学校排版。

## 何时用

- 会话里已有 / 刚写好的 `.md`，要变成可打开的 `.docx`
- 再接：润色（revise）、填表、或 `*-thesis-format` 套版

## 何时不用

- 已有 Word 来稿 → 直接 revise / form-fill / cupk
- 需要完美还原 MD（HTML、脚注、嵌套列表、图片）→ 本脚本不承诺；可先转再人工补

## Agent 硬流程

```text
1. 确认会话工作区中的 article.md（用户来稿只读时先复制到会话目录）
2. python <resources_path>/scripts/md_to_docx.py article.md --out article.docx --print
3. 需要学校格式 → 加载 *-thesis-format，把 article.docx 当「来稿」套版
4. 需要润色 → scenarios/revise.md
5. 禁止在未转 Word 前宣称「已按论文格式排好」
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

输出是**可编辑草稿**，中文字体默认宋体 + Times New Roman；页眉页码、石大格式由后续 skill 处理。
