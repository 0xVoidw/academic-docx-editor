"""Native Word Track Changes editing."""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

from .edit import plan_edits
from .models import EditorError, UnsafeEdit, ValidationError
from .ooxml import NS, Package, body_paragraphs, commit_new_file, copy_new_file, paragraph_text, sha256, write_new_text
from .word_session import check_lock, opened_document, word_application


def revise_document(input_path: Path, output_path: Path, changes: list[dict], dry_run: bool = False) -> dict:
    input_path, output_path = Path(input_path), Path(output_path)
    if input_path.resolve() == output_path.resolve() or output_path.exists():
        raise EditorError("Output must be a new path distinct from input")
    pkg = Package(input_path)
    plans = plan_edits(pkg, changes)
    body_text = [paragraph_text(p) for p in body_paragraphs(pkg.xml("word/document.xml"))]
    for plan in plans:
        if sum(text.count(plan["old"]) for text in body_text) != 1:
            raise UnsafeEdit("Revision mode requires an original string unique in the Word body")
        if any(plan["old"] in other["new"] for other in plans if other is not plan):
            raise UnsafeEdit("A replacement would create text matching a pending Word search")
    result = {"mode": "revisions", "input": str(input_path.resolve()),
              "output": str(output_path.resolve()), "dry_run": dry_run,
              "source_sha256": sha256(input_path), "changes": plans}
    if dry_run:
        return result
    check_lock(input_path)
    backup = output_path.with_name(output_path.name + ".source.bak.docx")
    manifest = output_path.with_name(output_path.name + ".manifest.json")
    if backup.exists() or manifest.exists():
        raise EditorError("Backup or manifest already exists")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    copy_new_file(input_path, backup)
    temporary = None
    committed = False
    try:
        with tempfile.NamedTemporaryFile(suffix=".docx", prefix=".academic-revisions-", dir=output_path.parent, delete=False) as tmp:
            temporary = Path(tmp.name)
        temporary.unlink()
        with word_application() as app:
            with opened_document(app, input_path, read_only=False) as document:
                document.TrackRevisions = True
                for plan in plans:
                    rng = document.Content
                    found = rng.Find.Execute(FindText=plan["old"], Forward=True, Wrap=0, MatchCase=True)
                    if not found or rng.Text != plan["old"]:
                        raise UnsafeEdit(f"Word could not locate exact text: {plan['old']}")
                    rng.Text = plan["new"]
                document.SaveAs2(FileName=str(temporary.resolve()), FileFormat=16, AddToRecentFiles=False)
        revised = Package(temporary)
        root = revised.xml("word/document.xml")
        insertions = len(root.findall(".//w:ins", NS))
        deletions = len(root.findall(".//w:del", NS))
        if insertions < len(plans) or deletions < len(plans):
            raise ValidationError("Word did not save expected native insertion/deletion revisions")
        from .validate import compare_revisions
        check = compare_revisions(input_path, temporary, plans)
        if not check["passed"]:
            raise ValidationError("Native revision verification failed: " + "; ".join(check["issues"]))
        commit_new_file(temporary, output_path)
        committed = True
        result.update({"backup": str(backup), "output_sha256": sha256(output_path),
                       "native_revisions": {"insertions": insertions, "deletions": deletions},
                       "validation": check})
        write_new_text(manifest, json.dumps(result, ensure_ascii=False, indent=2))
        result["manifest"] = str(manifest)
        return result
    except Exception:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
        if committed:
            output_path.unlink(missing_ok=True)
        raise
