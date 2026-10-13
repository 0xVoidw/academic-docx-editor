# Document safety

The CLI never writes to its input path or an existing output. A successful write produces `<output>.source.bak.docx` and `<output>.manifest.json`. The backup is a byte-for-byte copy of the original. If a write or post-write check fails, the partial output is removed; the backup remains available for recovery.

DOCX files are ZIP packages. Input validation checks ZIP CRC, duplicate members, XML parsing, required parts, and internal relationship targets. `validate` compares paragraph text, sections, tables, images, formulas, fields, hyperlinks, headers/footers, captions and OOXML part bytes. For plain edits, supply the edit manifest so only planned text changes are allowed. For image replacement, supply the manifest so exactly one image part may change.

Native Word revisions may rewrite OOXML packaging and run layout. Check the `w:ins` and `w:del` counts, open the result in Word, and accept/reject a sample revision manually before delivery. Do not claim byte-level structural equivalence for a Word-resaved document.

PDF validation checks page count, dimensions, renders every page when `--render-dir` is set, and computes changed-pixel fractions with configurable tolerance. A matching page count or low pixel difference does not prove correct academic content, equations, numbering, or line wrapping.

Exit codes: 0 success; 2 invalid input; 3 unsafe edit; 4 missing dependency; 5 filesystem/Word failure; 6 validation failure. Error output is JSON on stderr.
