import pytest

from academic_docx_editor.validate import compare_pdfs


@pytest.mark.pdf
def test_pdf_render_and_difference(tmp_path):
    fitz = pytest.importorskip('fitz')
    pytest.importorskip('PIL')
    before = tmp_path / 'before.pdf'
    after = tmp_path / 'after.pdf'
    for path, label in ((before, 'Before'), (after, 'After')):
        doc = fitz.open()
        page = doc.new_page(width=595, height=842)
        page.insert_text((72, 72), label)
        doc.save(path)
        doc.close()
    same = compare_pdfs(before, before, tmp_path / 'render-same')
    assert same['passed'] and len(same['rendered_pages']) == 2
    changed = compare_pdfs(before, after, tmp_path / 'render-changed', max_changed_fraction=0)
    assert not changed['passed']
    assert changed['changed_pixel_fractions'][0] > 0
