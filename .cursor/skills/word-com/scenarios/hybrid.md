# 场景：混合文档（表 + 分段框架）

上半 **表格填字段**，下面按 **标题** 填各段正文（申请书等）。

## 流程

```
已有申请书框架.docx + hybrid.json → hybrid_ops.py → Save
（可选）create_hybrid_template.py 生成空白框架供本地演示
```

## 命令

```bash
# 可选：生成空白模板
python <resources_path>/scripts/create_hybrid_template.py
# 必选：按 JSON 填写
python <resources_path>/scripts/hybrid_ops.py "申请书.docx" --json hybrid.json
```

## hybrid.json

```json
{
  "table": 1,
  "table_fields": {"姓名": "李四", "学院": "石油学院"},
  "sections": [
    {"heading": "第一部分  申请理由", "body": "……"},
    {"heading": "第二部分  个人情况", "body": "……"}
  ]
}
```

## 实现要点

- 表：`fill_by_fields`（与 fill_form 共用）
- 节：`find_paragraph_by_heading` 定位标题段，改下一段 `Range.Text`

## 交付

- `table_filled` / `sections_filled` / `missing_sections` 打印在 stdout
