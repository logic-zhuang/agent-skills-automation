# 填表用例：中国石油奖学金（WPS）

档案共用：[`../_shared/profile.txt`](../_shared/profile.txt)

## 一键

```bash
python wps-com/examples/form-fill/scholarship-cnpc/run_scholarship_demo.py
```

流程：`parse_profile` → `inspect` → `plan` → `auto_fill`（TrackRevisions，默认不定稿）→ 在 **WPS** 审阅。
