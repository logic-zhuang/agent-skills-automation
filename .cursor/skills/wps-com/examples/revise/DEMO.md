# 润色演示（WPS）

1. `create_short_essay.py` — 用 WPS COM 生成三段落中文短文
2. `demo_minimal_revision.py` — 对第 2 段做最小 diff 修订（TrackRevisions）

## 运行

```bash
python wps-com/examples/revise/demo_minimal_revision.py
```

## 预期

- stdout 含 `backend Kwps.Application` 或 `wps.Application`
- `revisions` > 0
- WPS 审阅面板可见删除/插入标记
- **不** AcceptAll
