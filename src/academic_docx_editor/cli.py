"""JSON-first command line interface."""

from __future__ import annotations

import argparse
import importlib.util
import json
import platform
import shutil
import sys
from pathlib import Path

from . import __version__
from .audit import inspect_docx, markdown_report
from .edit import edit_document
from .export import export_pdf
from .figures import bar_chart, flowchart
from .images import replace_inline_image
from .models import EditorError
from .ooxml import write_new_text
from .revisions import revise_document
from .validate import compare_documents, compare_pdfs, compare_revisions
from .word_session import word_application, word_executable


def doctor(probe_word: bool = False) -> dict:
    modules = {name: importlib.util.find_spec(name) is not None for name in
               ("docx", "lxml", "win32com", "fitz", "PIL", "matplotlib", "graphviz", "pytest")}
    word = word_executable()
    result = {"python": sys.version.split()[0], "python_supported": sys.version_info >= (3, 11),
              "platform": platform.platform(), "modules": modules,
              "word_executable": word, "word_registered": bool(word),
              "graphviz_dot": shutil.which("dot"), "word_com_probe": "not_run"}
    if probe_word:
        try:
            with word_application() as app:
                result["word_com_probe"] = {"ok": True, "version": app.Version}
        except EditorError as exc:
            result["word_com_probe"] = {"ok": False, "error": str(exc)}
    return result


def parser() -> argparse.ArgumentParser:
    top = argparse.ArgumentParser(prog="academic-docx-editor", description="Conservative academic DOCX editing")
    top.add_argument("--version", action="version", version=__version__)
    commands = top.add_subparsers(dest="command", required=True)
    d = commands.add_parser("doctor", help="Check Python, optional dependencies and Word")
    d.add_argument("--probe-word", action="store_true", help="Launch and close a private Word COM instance")
    a = commands.add_parser("audit", help="Read-only DOCX structure report")
    a.add_argument("--input", type=Path, required=True)
    a.add_argument("--json-out", type=Path)
    a.add_argument("--markdown-out", type=Path)
    for command in ("edit", "revisions"):
        p = commands.add_parser(command, help="Safe text edit" if command == "edit" else "Native Word Track Changes")
        p.add_argument("--input", type=Path, required=True)
        p.add_argument("--output", type=Path, required=True)
        p.add_argument("--changes", type=Path, required=True, help="JSON array of old/new and optional location/context")
        p.add_argument("--dry-run", action="store_true")
    image = commands.add_parser("replace-image", help="Replace one audited inline PNG/JPEG")
    image.add_argument("--input", type=Path, required=True)
    image.add_argument("--output", type=Path, required=True)
    image.add_argument("--image-id", required=True)
    image.add_argument("--replacement", type=Path, required=True)
    image.add_argument("--dry-run", action="store_true")
    export = commands.add_parser("export-pdf", help="Export DOCX using a private Word COM instance")
    export.add_argument("--input", type=Path, required=True)
    export.add_argument("--output", type=Path, required=True)
    export.add_argument("--dry-run", action="store_true")
    val = commands.add_parser("validate", help="Compare DOCX structure and optionally PDF rendering")
    val.add_argument("--before", type=Path, required=True)
    val.add_argument("--after", type=Path, required=True)
    val.add_argument("--manifest", type=Path)
    val.add_argument("--pdf-before", type=Path)
    val.add_argument("--pdf-after", type=Path)
    val.add_argument("--render-dir", type=Path)
    val.add_argument("--max-changed-fraction", type=float, default=0.01)
    val.add_argument("--pixel-tolerance", type=int, default=20)
    fb = commands.add_parser("figure-bar", help="Generate monochrome bar chart from JSON data")
    fb.add_argument("--data", type=Path, required=True)
    fb.add_argument("--output", type=Path, required=True)
    ff = commands.add_parser("figure-flowchart", help="Generate monochrome Graphviz figure from JSON relations")
    ff.add_argument("--spec", type=Path, required=True)
    ff.add_argument("--output", type=Path, required=True)
    return top


def _load_json(path: Path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise EditorError(f"Cannot read JSON {path}: {exc}") from exc


def run(args: argparse.Namespace) -> dict:
    if args.command == "doctor":
        return doctor(args.probe_word)
    if args.command == "audit":
        report = inspect_docx(args.input)
        if args.json_out:
            if args.json_out.exists():
                raise EditorError(f"Report exists: {args.json_out}")
            args.json_out.parent.mkdir(parents=True, exist_ok=True)
            write_new_text(args.json_out, json.dumps(report, ensure_ascii=False, indent=2))
        if args.markdown_out:
            if args.markdown_out.exists():
                raise EditorError(f"Report exists: {args.markdown_out}")
            args.markdown_out.parent.mkdir(parents=True, exist_ok=True)
            write_new_text(args.markdown_out, markdown_report(report))
        return report
    if args.command == "edit":
        return edit_document(args.input, args.output, _load_json(args.changes), args.dry_run)
    if args.command == "revisions":
        return revise_document(args.input, args.output, _load_json(args.changes), args.dry_run)
    if args.command == "replace-image":
        return replace_inline_image(args.input, args.output, args.image_id, args.replacement, args.dry_run)
    if args.command == "export-pdf":
        return export_pdf(args.input, args.output, args.dry_run)
    if args.command == "validate":
        manifest = _load_json(args.manifest) if args.manifest else {}
        expected_image = manifest.get("media_part")
        changes = manifest.get("changes") if manifest.get("mode") == "plain" else None
        docx_check = (compare_revisions(args.before, args.after, manifest["changes"])
                      if manifest.get("mode") == "revisions" else
                      compare_documents(args.before, args.after, changes, expected_image))
        result = {"docx": docx_check}
        if bool(args.pdf_before) != bool(args.pdf_after):
            raise EditorError("Provide both --pdf-before and --pdf-after")
        if args.pdf_before:
            result["pdf"] = compare_pdfs(args.pdf_before, args.pdf_after, args.render_dir,
                                         args.max_changed_fraction, args.pixel_tolerance)
        result["passed"] = all(value["passed"] for value in result.values())
        return result
    if args.command == "figure-bar":
        return bar_chart(_load_json(args.data), args.output)
    if args.command == "figure-flowchart":
        return flowchart(_load_json(args.spec), args.output)
    raise EditorError("Unknown command")


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    try:
        result = run(args)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        if args.command == "validate" and not result["passed"]:
            return 6
        return 0
    except EditorError as exc:
        print(json.dumps({"ok": False, "error": type(exc).__name__, "message": str(exc),
                          "exit_code": exc.exit_code}, ensure_ascii=False), file=sys.stderr)
        return exc.exit_code
    except OSError as exc:
        print(json.dumps({"ok": False, "error": "OSError", "message": str(exc),
                          "exit_code": 5}, ensure_ascii=False), file=sys.stderr)
        return 5


if __name__ == "__main__":
    raise SystemExit(main())
