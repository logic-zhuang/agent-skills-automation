# 知识库与档案规范（Excel 填表）

用户**不**在知识库里放单元格坐标、slots 或 fill.json。  
那些是系统内部产物。

## L1 用户档案（用户可见）

与 `word-com` 相同约定，例如：

```text
姓名：李俊杰
学号：2024216758
学院：文理学院
联系电话：13912345678
```

Agent 调用：

```bash
python <resources_path>/scripts/parse_profile.py profile.txt --out profile.json
```

生成会话工作区临时 `profile.json`。缺字段 → `missing` → Agent 追问，不要猜。

若会话已加载 `word-com`，也可用其 `parse_profile.py`；本 skill 自带一份，避免强依赖。

## L2 系统缓存（可选）

| 内容 | 说明 |
|------|------|
| `form_fingerprint` | inspect 算出的结构指纹 |
| 上次确认的 label↔address | 同表二次填写可复用 |

本期 excel-com **只输出指纹字段占位**，不实现跨会话 KB API。

## Agent 对用户话术

- 正确：「已根据你的基本信息填好草稿，请在 Excel 里核对后再定稿。」
- 错误：「请提供 fill.json / 告诉我 B2 填什么。」（除非用户主动要调试）
