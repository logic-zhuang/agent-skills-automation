# OpenWorker 常见：Markdown → docx（WPS 侧）

会话里写好 MD 后粗转成 `.docx`，再用 WPS 打开润色或填表。

## 一键

```bash
python wps-com/examples/md-to-docx/run_demo.py
```

输出：`out/sample.docx`（不依赖 WPS COM；打开审阅需 WPS）。

## 手动

```bash
python wps-com/scripts/md_to_docx.py wps-com/examples/md-to-docx/sample.md --out out.docx --print
```
