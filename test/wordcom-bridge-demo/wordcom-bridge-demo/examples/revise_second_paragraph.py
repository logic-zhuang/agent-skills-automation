"""示例：AI 只修改当前 Word 文档第 2 段，并保留 Word 原生修订。"""

from word_bridge import WordBridge, WordBridgeError


def ask_ai_to_revise(paragraph_text: str) -> tuple[str, str]:
    """这里用固定结果模拟大模型；实际项目中替换为 LLM 调用。"""
    revised_text = (
        "那天放学后，天空突然下起大雨。我没有带伞，只好站在校门口等待。"
        "这时，同学小林把自己的雨伞分给我一半，我们一起冒雨走回了家。"
    )
    return paragraph_text, revised_text


def main() -> None:
    try:
        bridge = WordBridge.connect_active()
        paragraph_number = 2
        paragraph = bridge.document.Paragraphs.Item(paragraph_number).Range
        original_text = str(paragraph.Text).rstrip("\r\x07")

        # 实际应用中：把 original_text 和修改要求一起发送给大模型。
        old_text, new_text = ask_ai_to_revise(original_text)
        if old_text != original_text:
            raise RuntimeError("AI 返回的原文与 Word 第 2 段不一致，已停止修改。")

        count = bridge.replace_text_in_paragraph(
            paragraph_number,
            old_text,
            new_text,
        )
        print(f"第 {paragraph_number} 段修改完成，产生 {count} 次原生修订。")
        print("请回到 Word 的“审阅”选项卡查看，并手动接受或拒绝。")
    except (WordBridgeError, ValueError, IndexError, RuntimeError) as exc:
        print(f"操作失败：{exc}")


if __name__ == "__main__":
    main()
