# 场景：批改 / 润色 / 最小修订

用户要**改已有 WPS 正文**（润色、批改、改几个字、改某段）时走本流程。  
COM 细则见 [minimal-revision.md](../minimal-revision.md)。

## 何时用

- 润色、批改、替换词句、改第 N 段
- LLM 返回整段新文字也要 **diff 写回**，不要整段 `Range.Text`

## 何时不用

- 论文套版 → 学校 `*-thesis-format` skill
- 只填表格几项 → [form-fill.md](form-fill.md)
- 新建空白文档 / 追加新段 → 可用 Range 规则，不必 TrackRevisions

## Agent 流程

```
1. com_session 或复用已打开 doc
2. 读 old_text（段/句/全文）
3. LLM → new_text
4. expected_old 校验（防串段）
5. minimal_revision.py 写回（TrackRevisions；默认 Save）
6. 提示：WPS 审阅 → 所有标记 → 接受/拒绝；未 AcceptAll
7. 若只需预览：--no-save；若要另存：--draft out.docx
```

## 命令

```bash
python <wps_com>/scripts/minimal_revision.py "doc.docx" --paragraph 2 --new "润色后的段落"
python <wps_com>/scripts/minimal_revision.py "doc.docx" --replace "旧" "新"
python <wps_com>/scripts/minimal_revision.py --active --replace "旧" "新"
python <wps_com>/scripts/minimal_revision.py "doc.docx" --job revise.json
python <wps_com>/scripts/minimal_revision.py "doc.docx" --paragraph 2 --new "…" --draft draft.docx
```

## 演示

```bash
python wps-com/examples/revise/demo_minimal_revision.py
```

说明见 [examples/revise/DEMO.md](../examples/revise/DEMO.md)。

## 交付说明

- 默认已 Save（修订标记仍在）；Accept/Reject 由用户在 WPS 完成
- **禁止** AcceptAll
- 声明：是否只改了指定段/句；`revisions` 数量可在 stdout 看到
