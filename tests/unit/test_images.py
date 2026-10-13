from academic_docx_editor.audit import inspect_docx
from academic_docx_editor.images import replace_inline_image
from academic_docx_editor.models import UnsafeEdit
from academic_docx_editor.ooxml import Package

from conftest import PNG


def test_replace_inline_preserves_xml(rich_docx, tmp_path):
    from PIL import Image
    replacement = tmp_path / 'new.png'
    Image.new('RGB', (4, 4), 'black').save(replacement)
    output = tmp_path / 'with-image.docx'
    result = replace_inline_image(rich_docx, output, 'image-1', replacement)
    assert result['validation']['passed']
    a, b = Package(rich_docx), Package(output)
    assert a.files['word/document.xml'] == b.files['word/document.xml']
    assert b.files['word/media/image1.png'] == replacement.read_bytes()
    assert inspect_docx(output)['counts']['images'] == 2


def test_floating_image_refused(rich_docx, tmp_path):
    replacement = tmp_path / 'new.png'
    replacement.write_bytes(PNG)
    try:
        replace_inline_image(rich_docx, tmp_path / 'out.docx', 'image-2', replacement, True)
    except UnsafeEdit:
        pass
    else:
        raise AssertionError('floating image was accepted')
