"""Document and PDF checks. Every check reports its actual evidence."""

from __future__ import annotations

import hashlib
from pathlib import Path

from .audit import inspect_docx
from .models import MissingDependency, ValidationError
from .ooxml import NS, Package


def compare_documents(before: Path, after: Path, expected_changes: list[dict] | None = None,
                      expected_image: str | None = None) -> dict:
    a, b = inspect_docx(before), inspect_docx(after)
    issues = []
    for key in ("sections", "tables", "images", "formulas", "fields", "hyperlinks"):
        if a["counts"][key] != b["counts"][key]:
            issues.append(f"{key} count changed: {a['counts'][key]} -> {b['counts'][key]}")
    if len(a["paragraphs"]) != len(b["paragraphs"]):
        issues.append("Paragraph count changed")
    if a["headers_footers"] != b["headers_footers"]:
        issues.append("Header/footer text changed")
    expected = [p["text"] for p in a["paragraphs"]]
    if expected_changes:
        for item in sorted(expected_changes, key=lambda x: (x["paragraph_index"], x["start"]), reverse=True):
            i, start, end = item["paragraph_index"], item["start"], item["end"]
            expected[i] = expected[i][:start] + item["new"] + expected[i][end:]
    actual = [p["text"] for p in b["paragraphs"]]
    expected_captions = [{"paragraph_index": c["paragraph_index"],
                          "text": expected[c["paragraph_index"]]} for c in a["captions"]]
    if expected_captions != b["captions"]:
        issues.append("Captions changed outside the plan or lost caption identity")
    if expected != actual:
        indices = [i for i, (x, y) in enumerate(zip(expected, actual)) if x != y]
        issues.append(f"Unexpected paragraph text difference at {indices[:20]}")
    pkg_a, pkg_b = Package(before), Package(after)
    if set(pkg_a.files) != set(pkg_b.files):
        issues.append("OOXML part list changed")
    allowed_parts = {"word/document.xml"}
    if expected_image:
        allowed_parts.add(expected_image)
    for part in set(pkg_a.files) & set(pkg_b.files):
        if part not in allowed_parts and pkg_a.files[part] != pkg_b.files[part]:
            issues.append(f"Unexpected OOXML part changed: {part}")
    if expected_changes:
        from .edit import apply_plans
        apply_plans(pkg_a, expected_changes)
        if pkg_a.files["word/document.xml"] != pkg_b.files["word/document.xml"]:
            issues.append("Document XML has changes outside the planned text nodes")
    elif not expected_image and pkg_a.files["word/document.xml"] != pkg_b.files["word/document.xml"]:
        issues.append("Document XML changed without an edit manifest")
    if expected_image:
        if pkg_a.files.get(expected_image) == pkg_b.files.get(expected_image):
            issues.append("Target image bytes did not change")
        if pkg_a.files.get("word/document.xml") != pkg_b.files.get("word/document.xml"):
            issues.append("Document XML changed during image replacement")
    return {"passed": not issues, "issues": issues, "before_sha256": a["sha256"],
            "after_sha256": b["sha256"], "before_counts": a["counts"], "after_counts": b["counts"],
            "manual_review": ["在 Word 中检查分页、行距和题注位置", "核查公式、交叉引用及参考文献语义",
                              "若使用修订，人工确认每条修订可接受或拒绝"]}


def compare_revisions(before: Path, after: Path, plans: list[dict]) -> dict:
    a, b = inspect_docx(before), inspect_docx(after)
    issues = []
    expected = [p["text"] for p in a["paragraphs"]]
    for item in sorted(plans, key=lambda x: (x["paragraph_index"], x["start"]), reverse=True):
        i, start, end = item["paragraph_index"], item["start"], item["end"]
        expected[i] = expected[i][:start] + item["new"] + expected[i][end:]
    actual = [p["text"] for p in b["paragraphs"]]
    if expected != actual:
        issues.append("Visible paragraph text differs from planned revisions")
    for key in ("tables", "images", "formulas"):
        if a["counts"][key] != b["counts"][key]:
            issues.append(f"{key} count changed")
    root = Package(after).xml("word/document.xml")
    ins = len(root.findall(".//w:ins", NS))
    dele = len(root.findall(".//w:del", NS))
    if ins < len(plans) or dele < len(plans):
        issues.append("Expected native insertion/deletion marks are missing")
    return {"passed": not issues, "issues": issues, "insertions": ins, "deletions": dele,
            "manual_review": ["在 Word 中逐条确认修订可接受或拒绝", "人工检查分页、题注及复杂对象"]}


def _pdf_info(path: Path) -> tuple[list[tuple[float, float]], list]:
    try:
        import fitz
    except ImportError as exc:
        raise MissingDependency("PyMuPDF is required: pip install 'academic-docx-editor[pdf]'") from exc
    try:
        pdf = fitz.open(path)
        if pdf.is_encrypted:
            raise ValidationError(f"Encrypted PDF: {path}")
        sizes = [(round(page.rect.width, 2), round(page.rect.height, 2)) for page in pdf]
        return sizes, [page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False) for page in pdf]
    except Exception as exc:
        if isinstance(exc, ValidationError):
            raise
        raise ValidationError(f"Cannot inspect PDF {path}: {exc}") from exc


def compare_pdfs(before: Path, after: Path, render_dir: Path | None = None,
                 max_changed_fraction: float = 0.01, pixel_tolerance: int = 20) -> dict:
    if not 0 <= max_changed_fraction <= 1 or not 0 <= pixel_tolerance <= 255:
        raise ValidationError("Invalid visual difference threshold")
    sizes_a, pix_a = _pdf_info(Path(before))
    sizes_b, pix_b = _pdf_info(Path(after))
    issues = []
    if len(sizes_a) != len(sizes_b):
        issues.append(f"PDF page count changed: {len(sizes_a)} -> {len(sizes_b)}")
    if sizes_a != sizes_b:
        issues.append("PDF page dimensions changed")
    rendered = []
    if render_dir:
        render_dir = Path(render_dir)
        render_dir.mkdir(parents=True, exist_ok=True)
        for label, pages in (("before", pix_a), ("after", pix_b)):
            for index, pix in enumerate(pages, 1):
                target = render_dir / f"{label}-{index:03d}.png"
                pix.save(target)
                rendered.append(str(target))
    try:
        from PIL import Image, ImageChops
    except ImportError as exc:
        raise MissingDependency("Pillow is required: pip install 'academic-docx-editor[pdf]'") from exc
    fractions = []
    for first, second in zip(pix_a, pix_b):
        if first.width != second.width or first.height != second.height:
            fractions.append(1.0)
            continue
        ia = Image.frombytes("RGB", (first.width, first.height), first.samples)
        ib = Image.frombytes("RGB", (second.width, second.height), second.samples)
        diff = ImageChops.difference(ia, ib).convert("L")
        hist = diff.histogram()
        changed = sum(hist[pixel_tolerance + 1:])
        fractions.append(changed / (first.width * first.height))
    flagged = [i + 1 for i, value in enumerate(fractions) if value > max_changed_fraction]
    if flagged:
        issues.append(f"Visual difference exceeds threshold on pages {flagged}")
    return {"passed": not issues, "issues": issues, "page_count_before": len(sizes_a),
            "page_count_after": len(sizes_b), "page_sizes_before": sizes_a,
            "page_sizes_after": sizes_b, "changed_pixel_fractions": fractions,
            "rendered_pages": rendered, "max_changed_fraction": max_changed_fraction,
            "pixel_tolerance": pixel_tolerance,
            "manual_review": ["逐页检查段落换行、页眉页脚、图表、题注和公式；像素差异只是筛查指标"]}
