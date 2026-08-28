# WPS Writer COM 对象模型参考

WPS 文字 COM 接口与 Microsoft Word VBA 高度兼容；对象名（`Documents`、`Range`、`Tables` 等）通常一致。本 skill **只挂接 WPS**，不回退 `Word.Application`。

集合一律 **从 1 计数**。VBA 枚举在 Python 里通常写成整数。

## ProgID 与挂接

| ProgID | 说明 |
|--------|------|
| `Kwps.Application` | WPS 文字（优先） |
| `wps.Application` | 部分新版 WPS 统一入口 |

挂接（见 [scripts/wps_session.py](scripts/wps_session.py)）：

```python
WPS_PROGIDS = ("Kwps.Application", "wps.Application")

# 已有 WPS 窗口
win32.GetActiveObject("Kwps.Application")

# 新开实例
win32.DispatchEx("Kwps.Application")
```

- 避免 `Dispatch("Kwps.Application")` 绑到坏掉的 `gen_py` 缓存
- **禁止**在本 skill 中尝试 `Word.Application`
- 同时安装 Word + WPS 时，本 skill 只连 WPS ProgID

`attach_wps()` 成功后 `backend_progid()` 返回实际 ProgID，供 `inspect_doc` / `com_session` 打印。

## Application

| 成员 | 说明 |
|------|------|
| `Visible` | `True` 显示窗口 |
| `DisplayAlerts` | `0` 抑制对话框 |
| `ScreenUpdating` | 批处理时可 `False` |
| `Documents` | 已打开文档 |
| `Quit()` | 退出 WPS；先 Close 未存文档 |
| `Version` | WPS 版本字符串 |

## Document / Range / Paragraph / Table

与 Word COM 相同约定（详见 [word-com/reference.md](../word-com/reference.md) 对象节）：

- 打开：`Documents.Open(FileName=..., ReadOnly=...)`
- 改字用 `Range`，段末 `\r`，单元格 `\x07`
- 改段正文：`rng.MoveEnd(1, -1)` 再 `rng.Text = ...`
- 保存：`Save()` / `SaveAs2(path)`

## WPS 限制与差异（冒烟后维护）

以下能力在 WPS 上**可能**与 Word 不同；脚本遇错应清晰报错，不假装成功：

| 能力 | 状态 | 说明 |
|------|------|------|
| `Kwps.Application` | 已验证 | 本机 WPS 12.0 |
| `wps.Application` | 未注册 | 本机仅 `Kwps.Application` 可用 |
| `TrackRevisions` / `Revisions` | 已验证 | 润色 demo 产生 6 条修订 |
| `inspect_doc` 已打开文档 | 已修复 | 按文件名回退匹配，避免重复 Open 失败 |
| `OMaths` / 公式 | 待验证 | 论文公式场景单独回归 |
| 页眉 `LinkToPrevious` | 大体兼容 | 复杂奇偶页眉各测一份 |
| 分节 `InsertBreak(2)` | 大体兼容 | |
| 样式 `NameLocal` | 大体兼容 | 中文 `"标题 1"` / `"正文"`；以 inspect 输出为准 |
| `.doc` 老格式 | 可打开 | 建议统一 `.docx` 再自动化 |
| 进程 `Quit` | 待验证 | 无文档时显式 Quit；避免 COM 泄漏 |

更新本表：在本机 WPS 上跑 `examples/revise/demo_minimal_revision.py` 与 `inspect_doc.py` 后记录结果。

## 常量（与 Word 共用 wd* 值）

| 名称 | 值 | 用途 |
|------|-----|------|
| `wdCharacter` | 1 | `MoveEnd` |
| `wdCollapseEnd` | 0 | `Collapse` |
| `wdSectionBreakNextPage` | 2 | `InsertBreak` |
| `wdStatisticPages` | 2 | `ComputeStatistics` |

## 故障排查

| 现象 | 处理 |
|------|------|
| `WPS Writer COM unavailable` | 安装 WPS Office；确认 ProgID 已注册 |
| 误连 Word | 本 skill 不应发生；检查是否加载了 word-com |
| `gen_py` / 奇怪 COM 错 | 用 `DispatchEx` / `GetActiveObject`；可删 `%TEMP%\gen_py` |
| 修订不可见 | 检查 WPS 审阅面板；确认 `TrackRevisions=True` |

## 与 word-com 文档

对象模型细节、Font/Section/Find 等长篇说明见 [word-com/reference.md](../word-com/reference.md)。WPS 自动化按同一 Range 规则编写；差异以本节 **WPS 限制** 为准。

**产品能力**（MD→docx / 润色 / 填表 / hybrid）与 word-com 对齐；`md_to_docx.py` 不依赖 COM，两边各有一份相同实现。
