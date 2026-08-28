# 交流生成绩录入（行式表 + append_row）

`layout_hint=table_rows`：Agent 用 `fill_add_course.json` 的 `append_row`，不用左→右 suggestions。

```bash
python excel-com/examples/form-fill/exchange-grades/run_demo.py
```

- 原文件只复制到 `_out/form.xlsx`
- 填写结果：`_out/draft_add_course.xlsx`
- OpenWorker 正式跑：同样 `--fill` + `--draft` 指到**会话工作区**
