"""Read-only DOCX inventory with stable paragraph and image locations."""

from __future__ import annotations

import re
from pathlib import Path

from .ooxml import NS, Package, body_paragraphs, paragraph_style, paragraph_text, part_relationships, q, sha256


def _text(node) -> str:
    return "".join(node.itertext()).strip()


def _properties(pkg: Package) -> dict:
    result = {}
    for part in ("docProps/core.xml", "docProps/app.xml"):
        if part in pkg.files:
            root = pkg.xml(part)
            result[part] = {child.tag.rsplit("}", 1)[-1]: _text(child) for child in root}
    return result


def inspect_docx(path: Path) -> dict:
    pkg = Package(path)
    root = pkg.xml("word/document.xml")
    paragraphs = body_paragraphs(root)
    styles = {}
    if "word/styles.xml" in pkg.files:
        for style in pkg.xml("word/styles.xml").findall("w:style", NS):
            sid = style.get(q("w", "styleId"))
            name = style.find("w:name", NS)
            styles[sid] = name.get(q("w", "val")) if name is not None else sid
    para_rows = []
    for index, p in enumerate(paragraphs):
        content = paragraph_text(p)
        sid = paragraph_style(p)
        para_rows.append({
            "index": index,
            "text": content,
            "sha256": __import__("hashlib").sha256(content.encode("utf-8")).hexdigest(),
            "style_id": sid,
            "style": styles.get(sid, sid),
            "run_count": len(p.findall(".//w:r", NS)),
            "in_table": p.getparent().tag != q("w", "body"),
        })
    sections = []
    for index, sect in enumerate(root.findall(".//w:sectPr", NS)):
        size = sect.find("w:pgSz", NS)
        margin = sect.find("w:pgMar", NS)
        sections.append({"index": index,
                         "page_twips": {k.rsplit("}", 1)[-1]: v for k, v in size.attrib.items()} if size is not None else None,
                         "margins_twips": {k.rsplit("}", 1)[-1]: v for k, v in margin.attrib.items()} if margin is not None else None})
    tables = []
    for index, table in enumerate(root.findall(".//w:tbl", NS)):
        rows = table.findall("w:tr", NS)
        tables.append({"index": index, "rows": len(rows), "cells_per_row": [len(r.findall("w:tc", NS)) for r in rows]})
    images = []
    captions = [{"paragraph_index": row["index"], "text": row["text"]} for row in para_rows
                if row["style_id"] in ("Caption", "caption") or re.match(r"^(图|表)\s*\d|^(Figure|Table)\s*\d", row["text"], re.I)]
    story_parts = [p for p in pkg.files if re.fullmatch(r"word/(header|footer)\d+\.xml", p)]
    stories = {part: [paragraph_text(p) for p in pkg.xml(part).iter(q("w", "p"))] for part in sorted(story_parts)}
    for part in ["word/document.xml", *sorted(story_parts)]:
        part_paragraphs = paragraphs if part == "word/document.xml" else list(pkg.xml(part).iter(q("w", "p")))
        rels = part_relationships(pkg, part)
        for pindex, p in enumerate(part_paragraphs):
            for drawing in p.findall(".//w:drawing", NS):
                kind = "floating" if drawing.find("wp:anchor", NS) is not None else "inline"
                extent = drawing.find(".//wp:extent", NS)
                docpr = drawing.find(".//wp:docPr", NS)
                for blip in drawing.findall(".//a:blip", NS):
                    rid = blip.get(q("r", "embed"))
                    item = rels.get(rid, {})
                    images.append({"id": f"image-{len(images)+1}", "part": part,
                                   "paragraph_index": pindex, "kind": kind,
                                   "relationship_id": rid, "media_part": item.get("target"),
                                   "name": docpr.get("name") if docpr is not None else None,
                                   "description": docpr.get("descr") if docpr is not None else None,
                                   "size_emu": {"cx": int(extent.get("cx")), "cy": int(extent.get("cy"))} if extent is not None else None})
    formulas = []
    fields = []
    cross_refs = []
    hyperlinks = []
    bookmarks = []
    for part in ["word/document.xml", *sorted(story_parts)]:
        tree = root if part == "word/document.xml" else pkg.xml(part)
        formulas.extend({"part": part, "type": "OMML"} for _ in tree.findall(".//m:oMath", NS))
        for fld in tree.findall(".//w:instrText", NS):
            code = fld.text or ""
            fields.append({"part": part, "code": code.strip()})
            if re.search(r"\b(REF|PAGEREF|SEQ)\b", code, re.I):
                cross_refs.append({"part": part, "code": code.strip()})
        for fld in tree.findall(".//w:fldSimple", NS):
            code = fld.get(q("w", "instr"), "")
            fields.append({"part": part, "code": code})
            if re.search(r"\b(REF|PAGEREF|SEQ)\b", code, re.I):
                cross_refs.append({"part": part, "code": code})
        part_rels = part_relationships(pkg, part)
        for link in tree.findall(".//w:hyperlink", NS):
            rid = link.get(q("r", "id"))
            hyperlinks.append({"part": part, "text": paragraph_text(link), "target": part_rels.get(rid, {}).get("target"), "anchor": link.get(q("w", "anchor"))})
        for mark in tree.findall(".//w:bookmarkStart", NS):
            bookmarks.append({"part": part, "name": mark.get(q("w", "name")), "id": mark.get(q("w", "id"))})
    anomalies = []
    for image in images:
        if not image["media_part"]:
            anomalies.append(f"{image['id']} has no resolvable media relationship")
        if image["kind"] == "floating":
            anomalies.append(f"{image['id']} is floating and protected from replacement")
    charts = len(root.findall(".//c:chart", NS))
    legacy_pictures = len(root.findall(".//w:pict", NS))
    embedded_objects = len(root.findall(".//w:object", NS))
    if charts:
        anomalies.append(f"{charts} embedded chart(s) require manual handling")
    if legacy_pictures or embedded_objects:
        anomalies.append(f"{legacy_pictures} legacy picture(s) and {embedded_objects} embedded object(s) require manual handling")
    try:
        from docx import Document
    except ImportError:
        secondary = {"status": "unavailable", "reason": "python-docx is not installed"}
    else:
        try:
            high_level = Document(path)
            secondary = {"status": "ok", "sections": len(high_level.sections), "tables": len(high_level.tables)}
            if secondary["sections"] != len(sections) or secondary["tables"] != len(tables):
                anomalies.append("python-docx and OOXML section/table counts disagree")
        except Exception as exc:
            secondary = {"status": "error", "reason": str(exc)}
            anomalies.append("python-docx could not open the document; verify in Word")
    return {"input": str(Path(path).resolve()), "sha256": sha256(Path(path)),
            "paragraphs": para_rows, "sections": sections, "tables": tables, "images": images,
            "captions": captions, "headers_footers": stories, "formulas": formulas,
            "fields": fields, "cross_references": cross_refs, "hyperlinks": hyperlinks,
            "bookmarks": bookmarks, "embedded_charts": charts, "legacy_pictures": legacy_pictures,
            "embedded_objects": embedded_objects, "properties": _properties(pkg),
            "python_docx_check": secondary,
            "anomalies": anomalies,
            "counts": {"paragraphs": len(para_rows), "sections": len(sections), "tables": len(tables),
                       "images": len(images), "formulas": len(formulas), "fields": len(fields),
                       "hyperlinks": len(hyperlinks)}}


def markdown_report(report: dict) -> str:
    counts = report["counts"]
    lines = ["# DOCX 审查报告", "", f"- 文件：`{report['input']}`", f"- SHA-256：`{report['sha256']}`", "",
             "## 结构概览", "", *[f"- {key}: {value}" for key, value in counts.items()], "",
             "## 标题与题注", ""]
    for p in report["paragraphs"]:
        style = (p["style"] or "").lower()
        if "heading" in style or "标题" in style or (p["style_id"] or "").startswith("Heading"):
            lines.append(f"- 段落 {p['index']} [{p['style']}]: {p['text']}")
    for caption in report["captions"]:
        lines.append(f"- 题注 {caption['paragraph_index']}: {caption['text']}")
    lines.extend(["", "## 图片", ""])
    lines.extend(f"- {im['id']}: 段落 {im['paragraph_index']}, {im['kind']}, {im['media_part']}" for im in report["images"])
    lines.extend(["", "## 需人工核查", ""])
    lines.extend(f"- {a}" for a in report["anomalies"])
    lines.extend(["- 核查公式、交叉引用和题注的语义与编号。", "- 用 Word 检查最终分页和修订显示。"])
    return "\n".join(lines) + "\n"
