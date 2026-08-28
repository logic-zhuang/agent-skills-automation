"""把《难忘的事》初稿写入当前活动 Word 文档。"""

import win32com.client


ESSAY = """难忘的事

每个人的成长过程中，都会遇到一些令人难忘的事情。它们有的让我们感到快乐，有的让我们懂得道理，还有的会在我们的心里留下温暖的回忆。对我来说，最难忘的是那次雨中的回家经历。

那天放学后，天空突然下起了倾盆大雨。我没有带伞，只好孤零零地站在校门口，焦急地望着雨幕。正当我不知道该怎么办的时候，同学小林撑着伞走了过来。他笑着对我说：“我们一起走吧！”于是，他主动把伞分给我一半。我们肩并肩走在雨中，虽然鞋子和裤脚都被雨水打湿了，但是我的心里却感到十分温暖。

到了第二天，我认真地向小林道了谢。他却摆摆手说：“同学之间本来就应该互相帮助。”这句话一直记在我的心里。以后遇到别人需要帮助时，我也会主动伸出援手。

这件小事虽然已经过去很久了，但每当下雨的时候，我都会想起那把小伞和同学真诚的笑容。
它让我明白，帮助别人不仅能给别人带来温暖，也能让自己的心灵变得更加美好。
这就是我最难忘的一件事。"""


def main() -> None:
    word = win32com.client.GetActiveObject("Word.Application")
    doc = word.ActiveDocument
    old_tracking = bool(doc.TrackRevisions)
    doc.TrackRevisions = False
    try:
        content = doc.Content
        existing = str(content.Text).rstrip("\r\x07")
        if existing.strip():
            content.InsertAfter("\r\r" + ESSAY)
        else:
            content.Text = ESSAY
    finally:
        doc.TrackRevisions = old_tracking
    word.Visible = True
    print("已将《难忘的事》写入当前活动 Word 文档。")
    print("初稿未生成修订；当前修订记录状态已恢复。")


if __name__ == "__main__":
    main()
