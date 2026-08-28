# word-com 测试用例

每个子目录是一套可独立跑的例子；以后往对应场景下加模板即可。

```text
examples/
  form-fill/
    _shared/                 # 共用学生档案（一张表也用这份）
      profile.txt
    scholarship-cnpc/        # 中国石油奖学金
    enrollment-undergrad/    # 本科新生入学登记表
    <你的新表>/
  revise/
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
