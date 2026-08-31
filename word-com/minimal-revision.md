# 最小修订（默认改稿策略）

**已有 Word 正文的修改，默认走本流程。** 场景入口见 [scenarios/revise.md](scenarios/revise.md)。用户不必说「最小修订」。

## 类 Git 的对比方式

| Git | Word + 本 skill |
|-----|-----------------|
| `git diff` | Word **审阅 → 所有标记**（删除线 + 下划线） |
| 只提交变更行 | `SequenceMatcher` 只写回**变了的字** |
| 人工 review 后 merge | 用户 **接受/拒绝** 每条修订 |
| 不自动 merge | **禁止** Agent 自动 `AcceptAll` |

LLM 可返回整段新文；写回前 diff，避免 Word 显示「整段删除 + 整段插入」。

## 默认 Save（连续 Agent 操作）

Agent 每步改完 **默认 `doc.Save()`**（`minimal_revision.py` 默认保存）。

原因：用户可能不在 Word 里点保存就发下一条 Agent 指令；不 Save 则下一步读盘仍是旧内容，`expected_old` 也会对不上。

- 默认：保存修订到 `.docx`（修订标记仍在，未 AcceptAll）
- 仅预览：脚本加 `--no-save`
- 另存草稿：`--draft out.docx`（不改原文件路径语义时用 SaveAs）
- 仍不自动 AcceptAll：Save ≠ 接受修订

## Agent 流程

```
1. 打开/复用文档
2. 读 old_text
3. LLM → new_text
4. expected_old 校验
5. minimal_revision.py 写回
6. doc.Save()（默认）
7. 提示：Word 审阅可对比；可继续下一步 Agent
```

## CLI

```bash
# 改第 N 段（默认 Save）
python <resources_path>/scripts/minimal_revision.py "essay.docx" --paragraph 2 --new "润色后的第2段"

# 改几个词
python <resources_path>/scripts/minimal_revision.py "essay.docx" --replace "只好" "只能"

# 活动 Word
python <resources_path>/scripts/minimal_revision.py --active --replace "旧" "新"

# 不写盘（少见）
python <resources_path>/scripts/minimal_revision.py "essay.docx" --paragraph 2 --new "…" --no-save

# 另存草稿
python <resources_path>/scripts/minimal_revision.py "essay.docx" --paragraph 2 --new "…" --draft draft.docx
```

### revise.json

```json
{
  "paragraph": 2,
  "old": "原文…",
  "new": "新文…",
  "expected_old": "原文…"
}
```

## 何时不用最小修订

- 新建文档 / 追加空段
- cupk 套版 `fill-section`
- 用户明确「直接覆盖，不要修订痕迹」

## 限制

- 表格、文本框、页眉页脚、OMath 内部：不要当普通全文 diff
- 变化极大时修订块仍可能较多
- 需要 Windows + 桌面 Word
