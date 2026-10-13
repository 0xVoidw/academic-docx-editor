import pytest
import shutil

from academic_docx_editor.audit import inspect_docx
from academic_docx_editor.cli import doctor
from academic_docx_editor.export import export_pdf
from academic_docx_editor.revisions import revise_document
from academic_docx_editor.validate import compare_revisions
from academic_docx_editor.word_session import opened_document, word_application


@pytest.fixture(scope='module')
def word_ready():
    probe = doctor(probe_word=True)
    if not isinstance(probe['word_com_probe'], dict) or not probe['word_com_probe']['ok']:
        pytest.skip(f"Word COM unavailable: {probe['word_com_probe']}")


@pytest.mark.word_com
def test_native_revisions(word_ready, simple_docx, tmp_path):
    output = tmp_path / 'tracked.docx'
    result = revise_document(simple_docx, output, [{'old': '研究标题', 'new': '修订标题'}])
    assert result['native_revisions']['insertions'] >= 1
    assert result['native_revisions']['deletions'] >= 1
    assert compare_revisions(simple_docx, output, result['changes'])['passed']
    with word_application() as app:
        with opened_document(app, output, read_only=True) as document:
            assert document.Revisions.Count >= 2
    accepted, rejected = tmp_path / 'accepted.docx', tmp_path / 'rejected.docx'
    shutil.copy2(output, accepted)
    shutil.copy2(output, rejected)
    for path, action in ((accepted, 'AcceptAll'), (rejected, 'RejectAll')):
        with word_application() as app:
            with opened_document(app, path, read_only=False) as document:
                document.TrackRevisions = False
                getattr(document.Revisions, action)()
                document.Save()
    assert inspect_docx(accepted)['paragraphs'][0]['text'].startswith('修订标题')
    assert inspect_docx(rejected)['paragraphs'][0]['text'].startswith('研究标题')


@pytest.mark.word_com
@pytest.mark.pdf
def test_word_pdf_export(word_ready, simple_docx, tmp_path):
    output = tmp_path / 'export.pdf'
    result = export_pdf(simple_docx, output)
    assert result['bytes'] > 100
