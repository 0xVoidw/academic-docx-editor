"""PDF export via isolated Microsoft Word COM."""

from __future__ import annotations

import tempfile
from pathlib import Path

from .models import EditorError, ValidationError
from .ooxml import Package, commit_new_file
from .word_session import check_lock, opened_document, word_application


def export_pdf(input_path: Path, output_path: Path, dry_run: bool = False) -> dict:
    input_path, output_path = Path(input_path), Path(output_path)
    Package(input_path)
    if output_path.suffix.lower() != ".pdf":
        raise EditorError("Output must have .pdf extension")
    if output_path.exists():
        raise EditorError(f"Output already exists: {output_path}")
    result = {"input": str(input_path.resolve()), "output": str(output_path.resolve()), "dry_run": dry_run}
    if dry_run:
        return result
    check_lock(input_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".pdf", prefix=".academic-pdf-", dir=output_path.parent, delete=False) as tmp:
            temporary = Path(tmp.name)
        temporary.unlink()
        with word_application() as app:
            with opened_document(app, input_path, read_only=True) as document:
                document.ExportAsFixedFormat(OutputFileName=str(temporary.resolve()), ExportFormat=17)
        if not temporary.is_file() or temporary.stat().st_size == 0:
            raise ValidationError("Word reported success but PDF is missing or empty")
        commit_new_file(temporary, output_path)
        result["bytes"] = output_path.stat().st_size
        return result
    except Exception:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
        raise
