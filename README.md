# academic-docx-editor v1.0

为中文学术论文及英文科技文档提供可审查的 DOCX 修改 CLI 和跨 Agent Skill。源码只有一份；Codex、Pi 与 Claude Code 通过目录链接发现同一个 `SKILL.md`。

## 安装

要求 Windows 11、Python 3.11+。基础安装包含 `lxml` 与 `python-docx`；Word 修订及 PDF 导出需要 Microsoft Word 和 `pywin32`；PDF 对比需要 PyMuPDF 与 Pillow；图表生成另需 Matplotlib、Graphviz Python 包和 `dot` 程序。以下 `py -3.11` 示例要求本机注册了相应 Python 版本；也可改用任一 3.11+ 的解释器完整路径。

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[all]"
.\.venv\Scripts\academic-docx-editor.exe doctor --probe-word
```

使用本仓库作为共享 Skill 源：

```powershell
.\scripts\install.ps1 -Agent All -Scope User -Action Install
.\scripts\install.ps1 -Agent All -Scope User -Action Verify
.\scripts\install.ps1 -Agent All -Scope User -Action Uninstall
```

项目级安装时加 `-Scope Project -ProjectRoot 'D:\your-project'`。Codex 与 Pi 共用 `.agents/skills` 链接；Claude 使用 `.claude/skills` 链接。已有文件或目录会被保留并报告冲突。仅当显式传入 `-UpdateLink` 时更新已有链接。早期 Codex 用户路径可加 `-CodexLegacy`。参见 [Agent 路径说明](references/agent-compatibility.md)。

## 工作流

```powershell
academic-docx-editor audit --input 'D:\论文\原稿.docx' --json-out audit.json --markdown-out audit.md
```

从审查 JSON 中获取 `paragraph_index` 和 `sha256`，创建 `changes.json`：

```json
[
  {
    "paragraph_index": 12,
    "paragraph_sha256": "填入 audit.json 对应段落的 sha256",
    "old": "本研究显示显著增加",
    "new": "本研究观察到增加",
    "before": "结果表明，",
    "after": "。"
  }
]
```

```powershell
academic-docx-editor edit --input 'D:\论文\原稿.docx' --output 'D:\论文\修订稿.docx' --changes changes.json --dry-run
academic-docx-editor edit --input 'D:\论文\原稿.docx' --output 'D:\论文\修订稿.docx' --changes changes.json
academic-docx-editor export-pdf --input 'D:\论文\原稿.docx' --output 'D:\论文\原稿.pdf'
academic-docx-editor export-pdf --input 'D:\论文\修订稿.docx' --output 'D:\论文\修订稿.pdf'
academic-docx-editor validate --before 'D:\论文\原稿.docx' --after 'D:\论文\修订稿.docx' --manifest 'D:\论文\修订稿.docx.manifest.json' --pdf-before 'D:\论文\原稿.pdf' --pdf-after 'D:\论文\修订稿.pdf' --render-dir 'D:\论文\逐页检查'
```

需要 Word 原生修订时，用 `revisions` 代替 `edit`。它另起 Word 实例并保存 `w:ins`/`w:del`；请在 Word 中确认每条修订可接受或拒绝。`replace-image` 使用审查报告的 `image-N` ID，仅支持唯一行内 PNG/JPEG 且保持原扩展名。`figure-bar` 和 `figure-flowchart` 从用户提供的数据/关系生成克制的黑白图。

所有写入命令拒绝覆盖输入或已有输出。普通修改和图片替换产生源文件备份及 manifest，并在落盘前检查 OOXML/结构差异。错误码与限制见 [安全说明](references/document-safety.md)。

## 支持矩阵

| 功能 | 状态 |
| --- | --- |
| DOCX 结构审查、JSON/Markdown 报告 | 已实现 |
| 唯一匹配、上下文和哈希定位、dry-run、纯文本安全修改 | 已实现 |
| 跨相同格式纯文本 Run | 已实现；复杂对象拒绝 |
| Word 原生修订、PDF 导出 | 已实现；Word 16.0 COM 集成测试通过 |
| 唯一行内 PNG/JPEG 替换 | 已实现；浮动图与嵌入图表仅识别保护 |
| PDF 逐页渲染与像素差异 | 已实现；已在可用的 PyMuPDF 环境验证 |
| 公式、域、超链接、嵌入 Excel 图表的自动改写 | 不支持，明确拒绝 |

运行测试：`python -m pytest -q`。COM 测试单独运行：`python -m pytest -m word_com -q`。测试使用自动生成的独立 DOCX，不读取用户论文。无 Word 或可选依赖时，相关测试跳过并由 `doctor` 报告实际状态。
