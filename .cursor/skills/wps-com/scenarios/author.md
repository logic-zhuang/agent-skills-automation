# 场景：OpenWorker 直接写 docx（默认写稿）

**默认**：用本 skill 的 WPS COM（`attach_wps` / 会话脚本）在会话工作区 **直接生成或写入 `.docx`**。  
**不要**先写 `.md` 再 `md_to_docx`，除非用户已有 Markdown 或明确要求转换。

## 何时用

- 「写一篇…」「生成 Word/WPS 文档」「起草…」等从零写稿
- 在已有 `.docx` 上追加新段

## 何时不用

- 用户**已经**给了 `.md` 且只要转成 docx → [md-to-docx.md](md-to-docx.md)
- 改已有正文 → [revise.md](revise.md)
- 只填表 → [form-fill.md](form-fill.md)

## Agent 硬流程

```text
1. 目标路径放在会话工作区；用户来稿只读
2. attach_wps + Documents.Add 或打开模板/空白 docx
3. Range / Paragraphs 写入（新建不必 TrackRevisions）
4. SaveAs2 / Save；Visible=True
5. 再改已有字 → revise
6. 禁止为写稿默认先 md 再转 docx
```

细则见 [reference.md](../reference.md)；与 [word-com/scenarios/author.md](../../word-com/scenarios/author.md) 对齐，仅 ProgID 不同。
