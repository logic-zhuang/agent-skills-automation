# 知识库与档案规范（OpenWorker 契约）

用户**不**在知识库里放表格坐标、slots、书签或 fill.json。  
那些是系统内部产物。

## L1 用户档案（用户可见）

用户通过表单或自然语言维护，例如：

- 姓名、学号、学院、专业、电话、身份证号
- 申请理由、成绩排名、是否综合考评等

可用键值文本：

```text
姓名：李俊杰
学号：2024216758
学院：文理学院
联系电话：13912345678
```

或一段自我介绍。Agent/App 调用：

```bash
python <resources_path>/scripts/parse_profile.py profile.txt --out profile.json
```

生成**会话工作区**临时 `profile.json`（内部用），不必让用户编辑。

缺字段时：`parse_profile` 返回 `missing` → Agent 追问，不要猜。

## L2 系统缓存（可选，系统写）

| 内容 | 说明 |
|------|------|
| `form_fingerprint` | inspect v2 算出的结构指纹 |
| 上次确认的 label↔fill_mode | 同表二次填写可复用 |

本期 word-com **只输出指纹**，不实现跨会话 KB API。OpenWorker 若实现 L2，用户应能「清除已记住的填表方式」。

## 明确不是知识库内容

- `slots.json` / 结构报告
- `fill.json` / plan 报告
- 书签映射、行列坐标表

## Agent 保底

1. 别名对齐（所在院系 ↔ college）
2. 档案缺字段 → 追问
3. **扫全看 `coverage`**（`unresolved_labels` / `unbound_empties`），不要用 `unmatched_kb` 判断漏扫
4. `needs_confirm` / `unknown` → 问整格还是拆格，或标请手填
5. 写入后请用户审阅；确认前不定稿、不 AcceptAll
