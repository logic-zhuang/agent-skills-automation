# 填表用例：中国石油奖学金申请审批表

本目录是一个**独立测试用例**（只放表与中间产物）。  
学生档案共用：[`../_shared/profile.txt`](../_shared/profile.txt)

## 文件

| 文件 | 作用 |
|------|------|
| `demo_scholarship_form.doc` | 空白表 |
| `demo_scholarship_draft.doc` | 带修订的草稿 |
| `demo_scholarship_slots.json` | inspect v2 |
| `demo_scholarship_plan.json` | plan |
| `run_scholarship_demo.py` | 一键演示 |

## 一键

```bash
python word-com/examples/form-fill/scholarship-cnpc/run_scholarship_demo.py
```

流程：`parse_profile` → `inspect` → `plan` → `auto_fill`（TrackRevisions，**不 --save**）→ 请你在 Word 里审阅。

确认后再定稿：

```bash
python word-com/scripts/auto_fill_form.py word-com/examples/form-fill/scholarship-cnpc/demo_scholarship_draft.doc --kb word-com/examples/form-fill/scholarship-cnpc/demo_profile.json --save
```

## 加新模板怎么放

```text
examples/form-fill/<短名>/
  README.md          # 说明表来源、注意点
  form.doc / .docx   # 空白原件
  profile.txt        # 测试用档案
  run_demo.py        # 可选：复制本目录脚本改路径
```

## 前提

Windows + Word + `pywin32`
