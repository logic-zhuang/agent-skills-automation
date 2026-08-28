# Word 原生修订示例

这个项目演示一个简单流程：

1. Python 生成一篇短作文并打开 Word。
2. 用户确认文章内容。
3. Python 模拟大模型修改第 2 段。
4. Word 使用原生“修订”显示删除和新增，用户最后接受或拒绝。

程序使用 Word 的 `TrackRevisions`，不是用红字或删除线模拟效果。

## 环境

- Windows
- 桌面版 Microsoft Word
- Python 3.11+
- [uv](https://docs.astral.sh/uv/)

安装依赖：

```text
uv sync
```

## 五分钟体验

### 第一步：生成并打开短作文

执行：

```text
uv run python examples/create_short_essay.py
```

脚本会创建并打开一篇只有 3 段的《难忘的事》：

```text
难忘的事

那天放学后突然下起大雨，我没有带伞，只好站在校门口等待。这时，同学小林把伞分给我一半，我们一起回家。

这件事让我感受到同学之间互相帮助的温暖。
```

如果 Word 没有自动显示，请打开生成的 `难忘的事.docx`，并确认它是当前活动文档。

### 第二步：修改第 2 段

执行：

```text
uv run python examples/revise_second_paragraph.py
```

这个脚本代表一次 AI 润色。它只读取第 2 段，并把它改成：

```text
那天放学后，天空突然乌云密布，紧接着下起了倾盆大雨。我没有带伞，只能焦急地站在校门口等待。这时，同学小林主动把雨伞分给我一半，我们一起冒雨回家。
```

脚本不会修改第 1 段和第 3 段。它会先比较新旧文字，只把变化部分交给 Word，因此相同的文字会保留，不会整段删除后再重新插入。

### 第三步：查看修订

回到 Word，打开“审阅”选项卡，将显示方式设置为“所有标记”。你可以看到：

- 原来的部分被标记为删除
- 新的部分被标记为插入
- 未变化的文字保持不变
- 可以点击“接受”或“拒绝”

程序不会关闭 Word，也不会自动保存文件。

## 三种修改方式

### 修改一句话

```python
from word_bridge import WordBridge

bridge = WordBridge.connect_active()
bridge.replace_text("旧句子", "新句子")
```

`replace_text()` 默认只修改全文中第一次出现的句子，并使用最小差异修订。

### 修改某一段

```python
paragraph = bridge.document.Paragraphs.Item(2).Range
old_text = str(paragraph.Text).rstrip("\r\x07")

bridge.replace_text_in_paragraph(
    2,
    old_text,
    "修改后的第 2 段",
)
```

段落编号从 `1` 开始。这个接口只在指定段落内查找，适合作文批阅。

### 修改全文

如果大模型重新生成了整篇文章，不要直接覆盖：

```python
doc.Content.Text = new_article
```

这样容易让 Word 显示整篇删除、整篇新增。应当保存修改前的全文，再使用：

```python
old_article = bridge.get_document_text()
new_article = ask_llm_to_rewrite(old_article)

bridge.replace_document_text(
    new_article,
    expected_old_text=old_article,
)
```

这个接口会比较新旧全文，只对变化部分产生修订。`expected_old_text` 用来检查 AI 生成期间 Word 是否被别人修改过；内容不一致时会停止写回，避免覆盖用户修改。

## 接入真实大模型

示例中目前用固定文本模拟 AI。真实流程只需要把模拟函数换成模型调用：

```text
读取 Word 第 2 段
    ↓
把段落和修改要求发送给大模型
    ↓
大模型返回修改后的段落
    ↓
校验原文没有变化
    ↓
调用 replace_text_in_paragraph()
    ↓
用户在 Word 中接受或拒绝
```

建议让模型返回：

```json
{
  "original_text": "Word 中原来的文字",
  "revised_text": "修改后的文字"
}
```

不要让模型直接操作 Word，也不要让模型自动接受全部修订。

## 示例文件

| 文件 | 用途 |
| --- | --- |
| `examples/create_short_essay.py` | 生成并打开短作文 |
| `examples/revise_second_paragraph.py` | 只修改第 2 段并产生修订 |
| `examples/write_essay.py` | 将较完整的《难忘的事》写入当前活动 Word |
| `examples/rewrite_revision_now.py` | 清除旧修订后重新生成局部修订 |
| `word_bridge/bridge.py` | Word COM 桥接核心代码 |

## 注意事项

- 必须先打开桌面版 Word，且目标文档是当前活动文档。
- 示例会直接改变 Word 内容，建议使用测试文档。
- 程序不会自动保存，也不会关闭 Word；确认结果后请手动保存。
- 文章变化很大时，Word 仍可能显示较多修订，这是正常的。
- 表格、图片、文本框、页眉页脚等复杂内容不适合直接当作普通纯文本全文替换。

## 测试

```text
uv run pytest
```
