---
name: academic-docx-editor
description: Audit and conservatively edit academic DOCX manuscripts, track changes in Microsoft Word, replace inline figures, export PDF, and verify structure/layout. Use for Chinese or English papers when original content and Word formatting must be protected.
metadata:
  version: "1.0.0"
---

# Academic DOCX Editor

Requires Python 3.11+; Word Track Changes and PDF export require Windows 11 with Microsoft Word and pywin32.

Use `academic-docx-editor` CLI for document operations. Start with `doctor` and `audit --input <docx> --json-out <report.json> --markdown-out <report.md>`. Read the report before proposing changes. The paragraph `index` and `sha256` are stable locators for a change file.

For text edits, prepare a JSON array with `old`, `new`, and preferably `paragraph_index`, `paragraph_sha256`, `before`, and `after`. Run `edit ... --dry-run` or `revisions ... --dry-run` first. Apply only after the preview uniquely resolves and the requested wording is approved by the user. Each write creates a new DOCX, a source backup, and a manifest; never target the original path.

Use `revisions` when native Word Track Changes is needed. Use `replace-image` only for an audited inline PNG/JPEG ID. Use `export-pdf` and `validate` after editing. Check the rendered PDF pages by eye; automated visual difference is only a screening signal.

Do not invent references, measurements, equations, or conclusions. Refuse complex Word objects the CLI cannot safely modify. If a command refuses a risky edit, tell the user what object blocked it and request a narrower edit or manual Word review.

- For safety rules and rollback, read [document-safety.md](references/document-safety.md).
- For text locator and revision behavior, read [editing-policy.md](references/editing-policy.md).
- For figure rules, read [academic-figures.md](references/academic-figures.md).
- For installation and agent discovery, read [agent-compatibility.md](references/agent-compatibility.md).
