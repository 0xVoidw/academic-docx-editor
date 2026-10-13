# Editing policy

Changes JSON example:

```json
[
  {
    "paragraph_index": 12,
    "paragraph_sha256": "SHA256_FROM_AUDIT",
    "old": "原始表述",
    "new": "经作者确认的新表述",
    "before": "前文",
    "after": "后文"
  }
]
```

Without a paragraph locator, `old` must occur exactly once in the entire body. With a locator, the paragraph hash must match the audit and the phrase must occur exactly once in that paragraph. Context is checked adjacent to the match. Overlapping changes are refused.

Plain edits modify only `w:t` nodes. A match within one node retains its run formatting. A match across text nodes is allowed only when all touched runs have identical run properties and no complex content is present. Hyperlinks, equations, field codes, drawings, tabs, breaks, nested tracked changes and mixed-format spans are refused. This conservative limit keeps format and OOXML objects intact; use a carefully reviewed Word edit for unsupported cases.

Revision mode uses a separately created Word COM instance (`DispatchEx`) and sets `TrackRevisions=True`. It only searches for preflighted unique body text and saves a separate DOCX. It does not attach to or quit an existing user Word process. Verify the native revisions in Word before final acceptance.
