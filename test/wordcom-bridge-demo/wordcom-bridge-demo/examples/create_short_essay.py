"""生成一篇短作文，保存并打开 Word 文档。"""

from pathlib import Path

import win32com.client


ESSAY = (
    "难忘的事\r"
    "那天放学后突然下起大雨，我没有带伞，只好站在校门口等待。"
    "这时，同学小林把伞分给我一半，我们一起回家。\r"
    "这件事让我感受到同学之间互相帮助的温暖。"
)


def main() -> None:
    word = win32com.client.Dispatch("Word.Application")
    document = word.Documents.Add()
    document.Content.Text = ESSAY
    output_path = Path.cwd() / "难忘的事.docx"
    document.SaveAs2(str(output_path), FileFormat=16)
    word.Visible = True
    document.Activate()
    print(f"已生成并打开：{output_path}")
    print("请确认文章内容后，再运行 revise_second_paragraph.py。")


if __name__ == "__main__":
    main()
