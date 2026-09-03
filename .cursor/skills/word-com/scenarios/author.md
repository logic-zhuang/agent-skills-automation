# 场景：OpenWorker 直接写 Word（默认写稿）

**默认**：用本 skill 的 Word COM（`attach_word` / 会话脚本）在会话工作区 **直接生成或写入 `.docx`**。  
**不要**先写 `.md` 再 `md_to_docx`，除非用户已有 Markdown 或明确要求「把 md 转成 word」。

## 何时用

- 「写一篇…」「生成 Word」「起草申请书正文」等从零写稿
- 在已有 `.docx` 上追加新段 / 新建空白文档填内容

## 何时不用

- 用户**已经**给了 `.md` 且只要转成 docx → [md-to-docx.md](md-to-docx.md)
- 改已有正文（润色、删字）→ [revise.md](revise.md)（必须 TrackRevisions）
- 只填表 → [form-fill.md](form-fill.md)

## Agent 硬流程

```text
1. 目标路径放在会话工作区（例 draft.docx）；用户来稿只读，勿原地覆盖
2. attach_word + Documents.Add 或打开模板/空白 docx
3. 用 Range / Paragraphs 写入标题与正文（新建内容不必开 TrackRevisions）
4. doc.SaveAs2 / Save；Visible=True，让用户在 Word 里看
5. 之后若再改已有字 → 转 revise（最小修订）
6. 禁止：为写稿先生成 article.md 再 md_to_docx（除非用户只要 md 或只要转换）
```

参考：`scripts/com_session.py`、`scripts/create_form_template.py` / `create_hybrid_template.py` 的 `Documents.Add` 写法；细则见 [reference.md](../reference.md)。

## 与 md-to-docx

| 情况 | 路径 |
|------|------|
| 默认写稿 | **本场景（COM）** |
| 用户已有 `.md` / 明确「转 Word」 | md-to-docx |
| 学校套版 | 已有 docx 后走 `*-thesis-format`（来稿是 md 时才先转） |
