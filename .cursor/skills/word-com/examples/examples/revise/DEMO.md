# 润色 / 最小修订演示

用一小段中文作文验证：LLM 可返回整段新文，但写回时只改变动字，Word 显示细粒度修订。

## 文件

| 文件 | 作用 |
|------|------|
| `create_short_essay.py` | 生成 `short_essay.docx` |
| `demo_minimal_revision.py` | 润色第 2 段（TrackRevisions + Save） |
| `revise_job.json` | 运行后生成的 job |
| `short_essay.docx` | 演示文档 |

## 一键

```bash
python word-com/examples/revise/create_short_essay.py
python word-com/examples/revise/demo_minimal_revision.py
```

或只跑第二步（没有作文时会自动创建）：

```bash
python word-com/examples/revise/demo_minimal_revision.py
```

## Agent 流程摘要

见 [scenarios/revise.md](../../scenarios/revise.md) 与 [minimal-revision.md](../../minimal-revision.md)。

要点：

- 默认 **Save**（方便连续改下一段）
- **不** AcceptAll；用户在 Word 审阅中接受/拒绝
- 仅预览用 `--no-save`；另存草稿用 `--draft out.docx`

## 手动命令示例

```bash
python word-com/scripts/minimal_revision.py word-com/examples/revise/short_essay.docx --paragraph 2 --new "润色后的第2段"
python word-com/scripts/minimal_revision.py word-com/examples/revise/short_essay.docx --replace "只好" "只能"
python word-com/scripts/minimal_revision.py --active --replace "旧" "新"
```
