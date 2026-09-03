# 场景：批改 / 润色 / 最小修订

用户要**改已有 Word 正文**（润色、批改、改几个字、改某段）时走本流程。  
COM 细则见 [minimal-revision.md](../minimal-revision.md)。

## 何时用

- 润色、批改、替换词句、改第 N 段
- **清理无用字符 / 多余空格 / 错符**（仍算改正文，必须出修订）
- LLM 返回整段新文字也要 **diff 写回**，不要整段 `Range.Text`

## 何时不用

- 论文套版 → 学校 `*-thesis-format` skill
- 只填表格几项 → [form-fill.md](form-fill.md)（填表另有 TrackRevisions；多选 □ 例外见 form-fill）
- 新建空白文档 / 追加新段 → 可用 Range 规则，不必 TrackRevisions
- 用户明确「不要修订痕迹、直接覆盖」

## Agent 流程

```
1. com_session 或复用已打开 doc
2. 读 old_text（段/句/全文）
3. LLM → new_text（删无用字时 new 更短即可）
4. expected_old 校验（防串段）
5. minimal_revision.py 写回（TrackRevisions + 强制所有标记；默认 Save）
   - 只删某串：--replace "无用串" ""
6. 核对 stdout 的 revisions>0；提示用户看删除线/下划线；未 AcceptAll
7. 若只需预览：--no-save；若要另存：--draft out.docx
```

**禁止**：关修订后 Delete；用 Selection/Backspace「悄悄删掉」。用户要看见改前（删除线）与改后。

## 命令

```bash
python <word_com>/scripts/minimal_revision.py "doc.docx" --paragraph 2 --new "润色后的段落"
python <word_com>/scripts/minimal_revision.py "doc.docx" --replace "旧" "新"
python <word_com>/scripts/minimal_revision.py --active --replace "旧" "新"
python <word_com>/scripts/minimal_revision.py "doc.docx" --job revise.json
python <word_com>/scripts/minimal_revision.py "doc.docx" --paragraph 2 --new "…" --draft draft.docx
```

## 交付说明

- 默认已 Save（修订标记仍在）；Accept/Reject 由用户在 Word 完成
- **禁止** AcceptAll
- 声明：是否只改了指定段/句；`revisions` 数量可在 stdout 看到
