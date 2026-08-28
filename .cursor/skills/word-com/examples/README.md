# word-com 测试用例

本地联调用；OpenWorker 正式填写输出在**会话工作区**。填表主路径：inspect `facts` → Agent `--fill`；旧表可用 `--auto-suggest`。

```text
examples/
  md-to-docx/                # OpenWorker MD → 粗 docx
  form-fill/
    _shared/                 # 共用学生档案
    scholarship-cnpc/        # --auto-suggest 回归
    enrollment-undergrad/    # --auto-suggest 回归
  revise/
```

## Markdown → Word

```bash
python word-com/examples/md-to-docx/run_demo.py
```

## 填表

```bash
python word-com/examples/form-fill/scholarship-cnpc/run_scholarship_demo.py
python word-com/examples/form-fill/enrollment-undergrad/run_demo.py
```

档案只改：`examples/form-fill/_shared/profile.txt`

## 润色

```bash
python word-com/examples/revise/create_short_essay.py
python word-com/examples/revise/demo_minimal_revision.py
```
