"""Replace an audited, unique inline raster image without changing layout XML."""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

from .audit import inspect_docx
from .models import EditorError, UnsafeEdit, ValidationError
from .ooxml import Package, commit_new_file, copy_new_file, part_relationships, sha256, write_new_text
from .validate import compare_documents
from .word_session import check_lock


def replace_inline_image(input_path: Path, output_path: Path, image_id: str,
                         replacement: Path, dry_run: bool = False) -> dict:
    input_path, output_path, replacement = Path(input_path), Path(output_path), Path(replacement)
    if input_path.resolve() == output_path.resolve() or output_path.exists():
        raise EditorError("Output must be a new path distinct from input")
    if not replacement.is_file():
        raise EditorError(f"Replacement image not found: {replacement}")
    check_lock(input_path)
    pkg = Package(input_path)
    audit = inspect_docx(input_path)
    image = next((item for item in audit["images"] if item["id"] == image_id), None)
    if not image:
        raise EditorError(f"Unknown image ID: {image_id}")
    if image["kind"] != "inline" or image["part"] != "word/document.xml":
        raise UnsafeEdit("Only inline images in the main document body can be replaced")
    part = image["media_part"]
    if not part or part not in pkg.files:
        raise UnsafeEdit("Image media relationship is unresolved")
    if replacement.suffix.lower() != Path(part).suffix.lower() or replacement.suffix.lower() not in (".png", ".jpg", ".jpeg"):
        raise UnsafeEdit("Replacement must be PNG/JPEG with the same extension as the original media")
    if sum(item["media_part"] == part for item in audit["images"]) != 1:
        raise UnsafeEdit("Media part is shared by multiple displayed images")
    references = sum(1 for owner in pkg.files if owner.endswith(".xml")
                     for rel in part_relationships(pkg, owner).values()
                     if rel["target"] == part and not rel["external"])
    if references != 1:
        raise UnsafeEdit("Media part is shared by another OOXML relationship")
    try:
        from PIL import Image
    except ImportError as exc:
        from .models import MissingDependency
        raise MissingDependency("Pillow is required for image verification") from exc
    try:
        with Image.open(replacement) as im:
            im.verify()
    except Exception as exc:
        raise ValidationError(f"Replacement is not a valid raster image: {exc}") from exc
    result = {"input": str(input_path.resolve()), "output": str(output_path.resolve()),
              "image_id": image_id, "media_part": part, "replacement": str(replacement.resolve()),
              "dry_run": dry_run, "source_sha256": sha256(input_path),
              "preserved_size_emu": image["size_emu"]}
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
        pkg.files[part] = replacement.read_bytes()
        with tempfile.NamedTemporaryFile(suffix=".docx", prefix=".academic-image-", dir=output_path.parent, delete=False) as tmp:
            temporary = Path(tmp.name)
        temporary.unlink()
        pkg.save(temporary)
        check = compare_documents(input_path, temporary, expected_image=part)
        if not check["passed"]:
            raise ValidationError("Image replacement verification failed: " + "; ".join(check["issues"]))
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
