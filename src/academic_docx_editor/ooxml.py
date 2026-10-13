"""Small, defensive OOXML layer shared by audit and mutation."""

from __future__ import annotations

import hashlib
import io
import posixpath
import shutil
import zipfile
from pathlib import Path

from lxml import etree

from .models import EditorError, ValidationError

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def q(prefix: str, local: str) -> str:
    return f"{{{NS[prefix]}}}{local}"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def xml_bytes(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


class Package:
    def __init__(self, path: Path):
        self.path = Path(path)
        if not self.path.is_file():
            raise EditorError(f"DOCX not found: {self.path}")
        try:
            with zipfile.ZipFile(self.path) as zf:
                if zf.testzip():
                    raise ValidationError("DOCX ZIP CRC check failed")
                names = zf.namelist()
                if len(names) != len(set(names)):
                    raise ValidationError("DOCX contains duplicate ZIP members")
                self.files = {name: zf.read(name) for name in names}
            if "word/document.xml" not in self.files or "[Content_Types].xml" not in self.files:
                raise ValidationError("Missing required DOCX package parts")
            for name, data in self.files.items():
                if name.endswith((".xml", ".rels")):
                    etree.fromstring(data, parser=etree.XMLParser(resolve_entities=False, no_network=True))
            self.verify_relationships()
        except (zipfile.BadZipFile, etree.XMLSyntaxError, OSError) as exc:
            raise ValidationError(f"Invalid DOCX: {exc}") from exc

    def xml(self, name: str) -> etree._Element:
        return etree.fromstring(self.files[name], parser=etree.XMLParser(resolve_entities=False, no_network=True))

    def verify_relationships(self) -> None:
        for name, data in self.files.items():
            if not name.endswith(".rels"):
                continue
            root = etree.fromstring(data)
            if name == "_rels/.rels":
                base = ""
            else:
                parent, relfile = posixpath.split(name)
                base = posixpath.dirname(parent)
            for rel in root:
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target", "")
                resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
                if resolved not in self.files:
                    raise ValidationError(f"Broken relationship: {name} -> {target}")

    def save(self, path: Path) -> None:
        path = Path(path)
        if path.resolve() == self.path.resolve():
            raise EditorError("Input file cannot be overwritten")
        if path.exists():
            raise EditorError(f"Output already exists: {path}")
        path.parent.mkdir(parents=True, exist_ok=True)
        stream = path.open("xb")
        try:
            with stream, zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as zf:
                for name, data in self.files.items():
                    zf.writestr(name, data)
        except Exception:
            path.unlink(missing_ok=True)
            raise


def copy_new_file(source: Path, output: Path) -> None:
    """Copy bytes to a new path only; clean a partial destination on failure."""
    with Path(source).open("rb") as input_stream:
        output_stream = Path(output).open("xb")
        try:
            with output_stream:
                shutil.copyfileobj(input_stream, output_stream, length=1024 * 1024)
        except Exception:
            Path(output).unlink(missing_ok=True)
            raise


def commit_new_file(temporary: Path, output: Path) -> None:
    copy_new_file(temporary, output)
    try:
        Path(temporary).unlink()
    except OSError:
        pass


def write_new_text(path: Path, content: str) -> None:
    stream = Path(path).open("x", encoding="utf-8")
    try:
        with stream:
            stream.write(content)
    except Exception:
        Path(path).unlink(missing_ok=True)
        raise


def body_paragraphs(root: etree._Element) -> list[etree._Element]:
    body = root.find("w:body", NS)
    return list(body.iter(q("w", "p"))) if body is not None else []


def paragraph_text(p: etree._Element) -> str:
    return "".join(n.text or "" for n in p.iter(q("w", "t")))


def paragraph_style(p: etree._Element) -> str | None:
    el = p.find("w:pPr/w:pStyle", NS)
    return el.get(q("w", "val")) if el is not None else None


def part_relationships(pkg: Package, part: str) -> dict[str, dict[str, str]]:
    directory, filename = posixpath.split(part)
    relpath = posixpath.join(directory, "_rels", filename + ".rels")
    if relpath not in pkg.files:
        return {}
    root = pkg.xml(relpath)
    result = {}
    for rel in root:
        target = rel.get("Target", "")
        resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(directory, target))
        result[rel.get("Id")] = {"type": rel.get("Type", ""), "target": resolved, "external": rel.get("TargetMode") == "External"}
    return result


def build_zip(files: dict[str, bytes]) -> bytes:
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    return stream.getvalue()
