# wps-com 测试用例

与 word-com **能力对齐**；COM 仅挂 WPS。OpenWorker 正式输出在**会话工作区**。

```text
examples/
  md-to-docx/                # MD → 粗 docx（无需 WPS COM）
  form-fill/
    _shared/                 # 共用学生档案
    scholarship-cnpc/
    enrollment-undergrad/
  revise/                    # 润色冒烟
```

## Markdown → docx

```bash
python wps-com/examples/md-to-docx/run_demo.py
```

## 填表

```bash
python wps-com/examples/form-fill/scholarship-cnpc/run_scholarship_demo.py
python wps-com/examples/form-fill/enrollment-undergrad/run_demo.py
```

档案只改：`examples/form-fill/_shared/profile.txt`  
空白表可与 word-com 同目录共用/复制。

## 润色（冒烟）

需已安装 WPS Office + pywin32：

```bash
python wps-com/examples/revise/create_short_essay.py
python wps-com/examples/revise/demo_minimal_revision.py
```

## 探查

```bash
python wps-com/scripts/inspect_doc.py wps-com/examples/revise/short_essay.docx
python wps-com/scripts/com_session.py wps-com/examples/revise/short_essay.docx
```

stdout 应含 `backend: Kwps.Application` 或 `wps.Application`。
