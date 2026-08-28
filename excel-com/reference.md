# Excel COM 参考（执行层）

## 应用与工作簿

```python
import win32com.client as win32

excel = win32.GetActiveObject("Excel.Application")  # 优先复用
# 或 win32.DispatchEx("Excel.Application")

excel.Visible = True
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(r"C:\path\form.xlsx")  # 仅当未打开时
ws = wb.Worksheets(1)  # 或 Worksheets("Sheet1")
```

复用已打开簿：遍历 `excel.Workbooks`，比较 `Path(wb.FullName).resolve()`。

## 单元格

| 写法 | 说明 |
|------|------|
| `ws.Range("B2")` | A1 地址 |
| `ws.Cells(2, 2)` | 行、列 1-based |
| `rng.Value2 = "张三"` | 写入标量 |
| `rng.Value2` | 读出；空多为 `None` |
| `rng.HasFormula` | 有公式则勿覆盖 |
| `rng.MergeCells` | 是否合并 |
| `rng.MergeArea.Cells(1, 1)` | 合并区左上角 |

## 使用区域

```python
used = ws.UsedRange
# used.Row, used.Column, used.Rows.Count, used.Columns.Count
```

空表 `UsedRange` 可能仍返回 A1；inspect 要容忍。

## 命名区域（可选）

```python
excel.Names.Item("姓名").RefersToRange.Value2 = "张三"
```

骨架期以「标签邻格」为主；命名区域可作为高置信 `fill_mode=named_range` 扩展。

## 草稿与定稿

| 行为 | API |
|------|-----|
| 另存草稿 | `wb.SaveAs(path)`（注意扩展名与 `FileFormat`） |
| 定稿 | `wb.Save()` |
| 保持打开 | **不要** `excel.Quit()` |

xlsx：`FileFormat=51`（`xlOpenXMLWorkbook`）。

## 审阅替代（无 TrackRevisions）

- 默认 `--draft out.xlsx`，原文件不动
- 可选：写入格 `Interior.Color = 0xCCFFCC`（浅色标记）
- plan 报告列出 `matched` / `needs_confirm` 给用户核对

## 常见坑

1. 同文件二次 `Open` → 冲突或只读副本；先 `find_open_workbook`。
2. 合并格写非左上角 → 静默失败或写错。
3. `DisplayAlerts=True` 时覆盖保存会弹窗阻塞 Agent。
4. 受保护工作表写入失败 → inspect/fill 应报错，勿静默跳过全部。
5. 日期/数字：字符串档案写入可能变成文本；业务表若需真日期，后续再扩类型。

## 常量速查

| 名 | 值 | 用途 |
|----|----|------|
| `xlOpenXMLWorkbook` | 51 | `.xlsx` SaveAs |
| `xlWorkbookDefault` | 51 | 同上（新版） |
| `xlUp` / `xlToLeft` 等 | -4162 / -4159 | End 导航（少用，优先 UsedRange） |
