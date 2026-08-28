"""清除旧修订，并对指定句子重新生成细粒度修订。"""

from word_bridge import WordBridge


OLD_TEXT = (
    "那天放学后，天空突然下起了倾盆大雨。"
    "我没有带伞，只好孤零零地站在校门口，焦急地望着雨幕。"
)
NEW_TEXT = (
    "那天放学后，天空突然乌云密布，紧接着下起了倾盆大雨。"
    "我没有带伞，只能孤零零地站在校门口，焦急地望着眼前的雨幕。"
)


def main() -> None:
    bridge = WordBridge.connect_active()
    bridge.reject_all_revisions()
    bridge.track_revisions = True
    count = bridge.replace_text(OLD_TEXT, NEW_TEXT)
    revisions = bridge.get_revisions()
    print(f"已清除旧修订，重新修改次数：{count}")
    print(f"当前原生修订数量：{len(revisions)}")
    for number, revision in enumerate(revisions, start=1):
        print(
            f"修订 {number}: 类型={revision.revision_type}, "
            f"内容={revision.text!r}"
        )
    print("请回到 Word，在“审阅-所有标记”中查看细粒度修订。")


if __name__ == "__main__":
    main()
