"""Preflight and conservative OOXML text editing."""

from __future__ import annotations

import hashlib
import json
import tempfile
from pathlib import Path

from lxml import etree

from .models import EditorError, UnsafeEdit, ValidationError
from .ooxml import NS, Package, body_paragraphs, commit_new_file, copy_new_file, paragraph_text, q, sha256, write_new_text, xml_bytes
from .word_session import check_lock


def _run_signature(run: etree._Element) -> bytes:
    props = run.find("w:rPr", NS)
    return etree.tostring(props, method="c14n") if props is not None else b""


def _simple_text_nodes(p: etree._Element) -> list[tuple[etree._Element, etree._Element, int, int]]:
    if p.findall(".//w:hyperlink", NS) or p.findall(".//m:oMath", NS) or p.findall(".//w:fldSimple", NS) or p.findall(".//w:instrText", NS) or p.findall(".//w:fldChar", NS):
        raise UnsafeEdit("Paragraph contains hyperlink, equation, or field code")
    nodes = []
    offset = 0
    for child in p:
        if child.tag == q("w", "pPr"):
            continue
        if child.tag != q("w", "r"):
            raise UnsafeEdit("Paragraph contains non-plain Word objects")
        for item in child:
            if item.tag not in (q("w", "rPr"), q("w", "t")):
                raise UnsafeEdit("Run contains tabs, drawings, breaks, or other non-text objects")
            if item.tag == q("w", "t"):
                length = len(item.text or "")
                nodes.append((item, child, offset, offset + length))
                offset += length
    return nodes


def plan_edits(pkg: Package, changes: list[dict]) -> list[dict]:
    if not isinstance(changes, list) or not changes:
        raise EditorError("Changes JSON must be a non-empty array")
    root = pkg.xml("word/document.xml")
    paras = body_paragraphs(root)
    original = [paragraph_text(p) for p in paras]
    plans = []
    used = set()
    for number, change in enumerate(changes, 1):
        old, new = change.get("old"), change.get("new")
        if not isinstance(old, str) or not old or not isinstance(new, str):
            raise EditorError(f"Change {number}: old must be nonempty text and new must be text")
        if old == new:
            raise EditorError(f"Change {number}: replacement is identical")
        positions = [(i, start) for i, text in enumerate(original) for start in _all_occurrences(text, old)]
        if "paragraph_index" in change:
            index = change["paragraph_index"]
            if not isinstance(index, int) or index < 0 or index >= len(paras):
                raise EditorError(f"Change {number}: invalid paragraph_index")
            positions = [(i, start) for i, start in positions if i == index]
            expected_hash = change.get("paragraph_sha256")
            if not expected_hash:
                raise UnsafeEdit(f"Change {number}: paragraph_sha256 required with paragraph_index")
            actual_hash = hashlib.sha256(original[index].encode("utf-8")).hexdigest()
            if expected_hash != actual_hash:
                raise UnsafeEdit(f"Change {number}: paragraph hash changed since audit")
        if len(positions) != 1:
            raise UnsafeEdit(f"Change {number}: expected one match, found {len(positions)}")
        index, start = positions[0]
        end = start + len(old)
        before, after = change.get("before"), change.get("after")
        if (before is not None and not isinstance(before, str)) or (after is not None and not isinstance(after, str)):
            raise EditorError(f"Change {number}: before/after context must be text")
        if before is not None and original[index][:start].endswith(before) is False:
            raise UnsafeEdit(f"Change {number}: before context mismatch")
        if after is not None and original[index][end:].startswith(after) is False:
            raise UnsafeEdit(f"Change {number}: after context mismatch")
        if any(i == index and not (end <= a or start >= b) for i, a, b in used):
            raise UnsafeEdit(f"Change {number}: overlaps another edit")
        nodes = _simple_text_nodes(paras[index])
        if not nodes or "".join(node.text or "" for node, _, _, _ in nodes) != original[index]:
            raise UnsafeEdit(f"Change {number}: paragraph text mapping is not safe")
        touched = [(node, run, a, b) for node, run, a, b in nodes if b > start and a < end]
        if not touched:
            raise UnsafeEdit(f"Change {number}: cannot resolve Word runs")
        if len({_run_signature(run) for _, run, _, _ in touched}) > 1:
            raise UnsafeEdit(f"Change {number}: match crosses runs with different formatting")
        used.add((index, start, end))
        plans.append({"paragraph_index": index, "start": start, "end": end,
                      "old": old, "new": new, "original_paragraph": original[index],
                      "paragraph_sha256": hashlib.sha256(original[index].encode("utf-8")).hexdigest(),
                      "run_count": len(touched)})
    return plans


def _all_occurrences(haystack: str, needle: str):
    start = 0
    while (at := haystack.find(needle, start)) != -1:
        yield at
        start = at + 1


def _set_text(node: etree._Element, value: str) -> None:
    node.text = value
    space_key = "{http://www.w3.org/XML/1998/namespace}space"
    if value.startswith(" ") or value.endswith(" "):
        node.set(space_key, "preserve")
    elif space_key in node.attrib:
        del node.attrib[space_key]


def apply_plans(pkg: Package, plans: list[dict]) -> None:
    root = pkg.xml("word/document.xml")
    paras = body_paragraphs(root)
    for plan in sorted(plans, key=lambda item: (item["paragraph_index"], item["start"]), reverse=True):
        p = paras[plan["paragraph_index"]]
        nodes = _simple_text_nodes(p)
        start, end = plan["start"], plan["end"]
        touched = [(node, a, b) for node, _, a, b in nodes if b > start and a < end]
        first, a, _ = touched[0]
        last, _, b = touched[-1]
        prefix = (first.text or "")[:start - a]
        suffix = (last.text or "")[end - (b - len(last.text or "")):]
        if first is last:
            _set_text(first, prefix + plan["new"] + suffix)
        else:
            _set_text(first, prefix + plan["new"])
            for node, _, _ in touched[1:-1]:
                _set_text(node, "")
            _set_text(last, suffix)
    pkg.files["word/document.xml"] = xml_bytes(root)


def edit_document(input_path: Path, output_path: Path, changes: list[dict], dry_run: bool = False) -> dict:
    input_path, output_path = Path(input_path), Path(output_path)
    if input_path.resolve() == output_path.resolve():
        raise EditorError("Input file cannot be overwritten")
    if output_path.exists():
        raise EditorError(f"Output already exists: {output_path}")
    check_lock(input_path)
    pkg = Package(input_path)
    plans = plan_edits(pkg, changes)
    result = {"mode": "plain", "input": str(input_path.resolve()), "output": str(output_path.resolve()),
              "dry_run": dry_run, "source_sha256": sha256(input_path), "changes": plans}
    if dry_run:
        return result
    backup = output_path.with_name(output_path.name + ".source.bak.docx")
    manifest = output_path.with_name(output_path.name + ".manifest.json")
    if backup.exists() or manifest.exists():
        raise EditorError("Backup or manifest already exists")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    copy_new_file(input_path, backup)
    temporary = None
    committed = False
    try:
        apply_plans(pkg, plans)
        with tempfile.NamedTemporaryFile(suffix=".docx", prefix=".academic-", dir=output_path.parent, delete=False) as tmp:
            temporary = Path(tmp.name)
        temporary.unlink()
        pkg.save(temporary)
        Package(temporary)
        from .validate import compare_documents
        check = compare_documents(input_path, temporary, plans)
        if not check["passed"]:
            raise ValidationError("Post-edit structure/text validation failed: " + "; ".join(check["issues"]))
        commit_new_file(temporary, output_path)
        committed = True
        result.update({"backup": str(backup), "output_sha256": sha256(output_path), "validation": check})
        write_new_text(manifest, json.dumps(result, ensure_ascii=False, indent=2))
        result["manifest"] = str(manifest)
        return result
    except Exception:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
        if committed:
            output_path.unlink(missing_ok=True)
        raise
