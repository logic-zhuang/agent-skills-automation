"""最小演示：修改当前打开 Word 中的第一处匹配文本。"""

from .bridge import WordBridge, WordBridgeError


def main() -> None:
    old = input("要查找的原文：")
    new = input("替换成：")
    try:
        bridge = WordBridge.connect_active()
        count = bridge.replace_text(old, new)
        print(f"完成：产生原生修订，替换 {count} 处。请回到 Word 查看并接受/拒绝。")
        for number, revision in enumerate(bridge.get_revisions(), start=1):
            print(
                f"修订 {number}: 类型={revision.revision_type}, "
                f"文本={revision.text!r}, 作者={revision.author}"
            )
    except (WordBridgeError, ValueError) as exc:
        print(f"操作失败：{exc}")


if __name__ == "__main__":
    main()
