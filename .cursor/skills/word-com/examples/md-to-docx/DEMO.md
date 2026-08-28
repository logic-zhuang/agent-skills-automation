# OpenWorker 常见：Markdown → Word

会话里写好 MD 后，用本用例粗转成 `.docx`，再交给润色或学校套版。

## 一键

```bash
python word-com/examples/md-to-docx/run_demo.py
```

输出：`out/sample.docx`

## 手动

```bash
python word-com/scripts/md_to_docx.py word-com/examples/md-to-docx/sample.md --out out.docx --print
```
