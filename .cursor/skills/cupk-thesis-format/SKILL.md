---
name: CUPK本科论文套版
description: 将本科论文 Word（.doc/.docx）转为中国石油大学（北京）克拉玛依校区 2023 格式：复制可填写模板，填充封面、中英文摘要与正文，处理图表题注、OMath、页眉页码。用户要求论文套版、石大/克拉玛依毕设格式，或附上待排版论文时使用。
source: repo
---

# CUPK 本科论文套版

把用户来稿变成克拉玛依校区 2023 本科格式。版式见 [rules.md](rules.md)；逐步操作见 [convert.md](convert.md)。Word COM 遵守 [Word文档操作](../word-com/SKILL.md)。

`resources_path` 来自 `load_skill("CUPK本科论文套版")`（本目录；技能目录名仍为 `cupk-thesis-format`）。脚本路径用 `/`。Shell cwd = OpenWorker 会话工作区。勿 `cd` 到 `resources_path`；`list_files` 若报超出范围则改用 shell。

## 路径

| 资源 | 位置 |
|------|------|
| 可填写底板 | `templates/…标准模板（可填写）.docx`（**`clone` 唯一默认源**；须随 skill 分发） |
| JSON 夹具 | `examples/`（字段结构参考；运行时 JSON 写在会话工作区） |
| 共享 COM | `scripts/cupk_com.py` |
| 权威范本 | 仅 **§模板维护** 重建底板时用；运行时不可回退 |

## 默认策略（必守）

**不要在用户原稿上直接改格式。** 复制可填写模板 → 填入来稿内容 → 删掉【格式提示】。

页码是范本里的 `PAGE` 域。用 `Fields.Add` 重写页脚会变成 `-  -1`。易错步骤用 `scripts/convert_ops.py`，不要手写清页脚。

## 套版流程

```
- [ ] python <resources_path>/scripts/inspect_source.py "来稿.docx" --extract extract.json
- [ ] 从 extract/来稿抽出封面六项；丢弃硕士字段（中图分类号、UDC、密级、学号、二导…）
- [ ] python <resources_path>/scripts/convert_ops.py clone --out "输出_CUPK本科格式_题目_姓名.docx"
- [ ] convert_ops.py cover --doc OUT.docx --json cover.json
- [ ] 按节 fill-section：摘要(2, --keep-heading + keywords) → ABSTRACT(3, --keep-paras 3 + title_en)
      → 勿动第4节目录 → 前言(5) → 第n章(6…) → 结论 → 参考文献 → 附录/致谢
- [ ] 章少：从后往前 delete-section 多余章；章多：extra-chapter --before 当前结论节号
- [ ] 图/表/公式：blocks 里 type=picture|omath，或 COM 粘贴；表尽量三线
- [ ] convert_ops.py finish --doc OUT.docx
- [ ] 抽查封面、摘要页码 `- I -`、正文 `- 1 -`、奇数页眉校名
```

## 必须 COM

- 封面只改 `Tables(1)` 文字，保留仿宋三号加粗；校徽/WordArt 各段居中
- **「标题 1」样式定义是两端对齐**：写章标题后必须再居中（用 `apply_heading1`；`finish` 会扫全文）
- 页眉 `LinkToPrevious = False` 后只改偶数页眉；**不要清空再插 PAGE 域**
- OMath：`OMaths.Add` → `finish` 里 `BuildUp()`；不要批量 `OMaths.Range.Delete`
- 同一路径禁止二次 `Open`

## 交付

输出 `.docx`。声明：三线表是否抽查、公式显示是否抽查、英文题注是否代译、目录是否已更新域。
