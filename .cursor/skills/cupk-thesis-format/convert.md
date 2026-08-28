# 来稿 Word → CUPK 套版细则

主流程见 [SKILL.md](SKILL.md)。版式数字见 [rules.md](rules.md)。

脚本均在 `<resources_path>/scripts/`。`resources_path` 来自
`load_skill("CUPK本科论文套版")`；shell cwd 保持 OpenWorker 会话工作区。

## 1. 探查来稿

```bash
python <resources_path>/scripts/inspect_source.py "path/to/thesis.docx"
python <resources_path>/scripts/inspect_source.py "path/to/thesis.docx" --extract extract.json
```

记下封面字段、摘要/关键词、英文题、章标题、表/图/公式数量。跳过 `~$*.docx`。

来稿常见形态：

| 形态 | 做法 |
|------|------|
| 已是本科结构（摘要→章→结论） | 按节映射到模板 |
| 硕士封面 | 只取题目/学院/专业/姓名/导师/起止时间 |
| PDF 转 Word / 纯文本 | 图可能是整页截图；表可能是文字块，交付时声明 |
| 无英文摘要/图题 | 代译并在交付说明列出 |

`extract.json` 的 `chapters[].blocks` 可改成 fill-section 用的 JSON（补图路径、删空段、统一「第n章  标题」空两格）。

## 2. 克隆底板

默认源为 `<resources_path>/templates/…标准模板（可填写）.docx`。缺失时 `clone` 报错，
见 [§模板维护](#模板维护)。

```bash
python <resources_path>/scripts/convert_ops.py clone --out "输出_CUPK本科格式_题目_姓名.docx"
```

用 word-com 打开输出文件，不要 Open 范本当工作稿。已存在则加 `--force`。可用 `--template` 指定其它底板。

## 模板维护

**运行时只读 `templates/`**；不要在 `cupk_com.fillable_template_path()` 里回退到
仓库其它目录。

从权威材料重建可填写底板（维护者操作，非套版必需步骤）：

1. 权威范本：`大学论文模板/中国石油大学（北京）克拉玛依校区/…格式范本（2023年修订）.docx`
   （开发仓库内；OpenWorker 分发包中可能没有）
2. 按 `rules.md` 封面/节/页眉页脚规则删提示、保留域，产出可填写版
3. 写入 `<resources_path>/templates/中国石油大学（北京）本科毕业设计（论文）标准模板（可填写）.docx`
4. 若有 `scripts/build_standard_template.py`，可运行：
   `python <resources_path>/scripts/build_standard_template.py`

套版前确认 `templates/` 下文件存在；否则 `clone` 会提示执行上述步骤。

## 3. 封面

`cover.json`（字段结构见 `examples/cover.json`）：

```json
{
  "title": "论文题目",
  "title_cont": "",
  "college": "石油学院",
  "major": "石油工程",
  "name": "张明轩",
  "advisor": "李建国 副教授",
  "date": "起止时间：2025 年 9 月 1 日 至 2026 年 6 月 15 日"
}
```

```bash
python <resources_path>/scripts/convert_ops.py cover --doc OUT.docx --json cover.json
```

只改单元格文字：标签列分散对齐、仿宋_GB2312 三号加粗；题目/续行居中。不续行则 `title_cont` 为空。不要恢复单位代码/学号文本框。

## 4. 往模板各节填内容

可填写模板 **14 节**（不要打乱页脚）：

| 节 | 内容 | fill-section |
|----|------|----------------|
| 1 | 封面 | `cover` 子命令 |
| 2 | `摘    要` + 正文 + `关键词：` | `--section 2 --keep-heading`，JSON 含 `keywords` |
| 3 | 英文题目 + 空行 + `ABSTRACT` + 正文 + `Key Words：` | `--section 3 --keep-paras 3`，JSON 含 `title_en`、`key_words` |
| 4 | `目    录` + TOC 域 | **不要 fill**；`finish` 会 `Update()` |
| 5 | `前    言` | `--section 5 --heading "前    言"` |
| 6–10 | 第1–5章 | `--heading "第n章  标题"`（空两格） |
| 11 | 结论 | `--heading "结    论"` 或 `第n章  结    论` |
| 12 | 参考文献 | 正文五号、悬挂缩进可事后抽查 |
| 13 | 附录 | 无附录：`delete-section --index 13` |
| 14 | 致谢 | 不可省 |

`ch1.json` 示例（夹具见 `examples/ch1.json`；图片路径用会话工作区或绝对路径）：

```json
{
  "heading": "第1章  绪论",
  "even_header": "第1章  绪论",
  "blocks": [
    {"style": "标题 2", "text": "1.1  研究目的与意义"},
    {"style": "正文", "text": "……"},
    {"type": "picture", "path": "fig.png", "caption_zh": "图1.1  题名", "caption_en": "Fig. 1.1  Title"},
    {"type": "omath", "linear": "k=a\\phi^b/S_wi^c", "number": "(1.1)"}
  ]
}
```

```bash
python <resources_path>/scripts/convert_ops.py fill-section --doc OUT.docx --section 6 --json ch1.json
```

摘要节：

```bash
python <resources_path>/scripts/convert_ops.py fill-section --doc OUT.docx --section 2 --keep-heading --json abs_zh.json
```

`abs_zh.json`：`{"blocks":[{"style":"正文","text":"……"}],"keywords":"词1；词2；词3"}`。

ABSTRACT 节：JSON 含 `title_en`、`blocks`、`key_words`。

插入位置由脚本在分节符前 `InsertBefore`。改已有段：`MoveEnd(1,-1)` 再赋 `Text`。

### 关键词

「关键词：」四字黑体 **小三（15）** 加粗；其后词黑体 **小四（12）** 加粗；分号分隔、末尾无标点。`finish` 会再扫一遍。Key Words 整行 Times New Roman 小四加粗。

### 样式

- 章：`标题 1`（`第n章` + 空两格 + 标题）。样式定义可能是两端对齐，**写完后必须再设居中**（`apply_heading1` / `finish` 会扫一遍）
- 节/小节：`标题 2` / `标题 3`，编号与标题空两格（`1.1  标题`）
- 其余：`正文`（JSON 文本已是纯文本，不要从原稿带样式粘贴）
- 标题颜色自动，不要主题蓝

### 章数与模板不一致

模板正文固定 5 个章节能（节 6–10）+ 结论（11）。

- **来稿章少**：从后往前删多余章，避免节号错位：

```bash
python <resources_path>/scripts/convert_ops.py delete-section --doc OUT.docx --index 10
python <resources_path>/scripts/convert_ops.py delete-section --doc OUT.docx --index 9
```

- **来稿章多**：在**当前结论节**之前插节。每插一章结论节号 +1：

```bash
python <resources_path>/scripts/convert_ops.py extra-chapter --doc OUT.docx --before 11 --json ch6.json
```

`extra-chapter`：`InsertBreak(2)` → 页眉页脚 `LinkToPrevious=False` → **页脚 Range 不 Delete、不 Fields.Add** → 奇数页眉校名，偶数页眉=本章标题。

### 图 / 表 / 公式

- 图：JSON `type=picture`；中文图题在图下黑体五号居中 `图2.1  …`；下一行 `Fig. 2.1  …` TNR 五号
- 表：Word 表，尽量三线（顶/底/表头下 0.5 磅）；中文表题在表上；`Table x.x`；表内宋体五号。从来稿 `Tables(i).Range.Copy()` 贴到目标节
- 公式：`type=omath`；编号 `(2.1)` 在公式外；`finish` 对 OMath `BuildUp()`；失败则保留线性式并在交付说明列出

## 5. 收口

```bash
python <resources_path>/scripts/convert_ops.py finish --doc OUT.docx
```

会删除 `【格式提示】`、重套关键词字号、把页眉里的「硕士学位论文」改成本科、BuildUp 公式、更新目录域。

抽查：

抽查时加载 [word-com](../word-com/SKILL.md)，取其 `resources_path` 为
`<word_com_resources_path>`：

```bash
python <word_com_resources_path>/scripts/inspect_doc.py "输出.docx" --start 1 --count 40
```

- 封面表有题目/姓名；WordArt 居中
- 第 2 节页脚奇数 `- I -`（宋体），结果为罗马数字
- 前言节起奇数页脚 `- 1 -`
- `Comments=0`；无单位代码文本框
- 无【格式提示】

## 6. 不要做

- 用 `Fields.Add(..., 33)` 重做页码
- 把范本「正文格式说明 / 图表及公式的格式说明 / 打印装订」拷进学生稿
- 改 `Application.UserName`、并行两个脚本打同一 docx
- 对已打开文件再 `Documents.Open` 一次
- 在用户原稿上直接套样式交差
