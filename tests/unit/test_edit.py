import hashlib

import pytest

from academic_docx_editor.audit import inspect_docx
from academic_docx_editor.edit import edit_document
from academic_docx_editor.models import EditorError, UnsafeEdit
from academic_docx_editor.ooxml import Package


def test_dry_run_and_cross_run_same_format(simple_docx, tmp_path):
    report = inspect_docx(simple_docx)
    paragraph = report['paragraphs'][1]
    change = [{'paragraph_index': 1, 'paragraph_sha256': paragraph['sha256'],
               'old': '观察到增加', 'new': '观察到轻微增加', 'before': '本研究', 'after': '。'}]
    output = tmp_path / '修订稿.docx'
    preview = edit_document(simple_docx, output, change, dry_run=True)
    assert preview['changes'][0]['run_count'] == 2
    assert not output.exists()
    result = edit_document(simple_docx, output, change)
    assert output.exists() and (tmp_path / '修订稿.docx.source.bak.docx').exists()
    assert result['validation']['passed']
    assert inspect_docx(output)['paragraphs'][1]['text'] == '本研究观察到轻微增加。'
    assert Package(simple_docx).files['word/document.xml'] != Package(output).files['word/document.xml']


def test_mixed_format_cross_run_rejected(simple_docx, tmp_path):
    with pytest.raises(UnsafeEdit, match='different formatting'):
        edit_document(simple_docx, tmp_path / 'out.docx', [{'old': 'Bold 斜体', 'new': 'changed'}], True)


def test_duplicate_match_rejected(simple_docx, tmp_path):
    with pytest.raises(UnsafeEdit, match='found 2'):
        edit_document(simple_docx, tmp_path / 'out.docx', [{'old': '重复', 'new': '唯一'}], True)


def test_input_cannot_be_overwritten(simple_docx):
    with pytest.raises(EditorError):
        edit_document(simple_docx, simple_docx, [{'old': 'Title', 'new': '标题'}], True)


def test_hyperlink_edit_refused(rich_docx, tmp_path):
    with pytest.raises(UnsafeEdit, match='hyperlink'):
        edit_document(rich_docx, tmp_path / 'out.docx', [{'old': '链接文本', 'new': '改写'}], True)


def test_failed_validation_rolls_back_output(simple_docx, tmp_path, monkeypatch):
    import academic_docx_editor.validate as validate
    monkeypatch.setattr(validate, 'compare_documents', lambda *args, **kwargs: {'passed': False, 'issues': ['injected failure']})
    output = tmp_path / 'rolled-back.docx'
    with pytest.raises(Exception, match='Post-edit structure'):
        edit_document(simple_docx, output, [{'old': 'Title', 'new': 'Heading'}])
    assert not output.exists()
    assert (tmp_path / 'rolled-back.docx.source.bak.docx').exists()


def test_caption_text_edit_preserves_caption_style(simple_docx, tmp_path):
    output = tmp_path / 'caption.docx'
    result = edit_document(simple_docx, output, [{'old': '表 1 示例', 'new': '表 1 研究样本'}])
    assert result['validation']['passed']
    assert inspect_docx(output)['captions'][0]['text'] == '表 1 研究样本'


def test_locked_source_rejected(simple_docx, tmp_path):
    lock = simple_docx.with_name('~$' + simple_docx.name[2:])
    lock.write_text('occupied')
    with pytest.raises(UnsafeEdit, match='lock file'):
        edit_document(simple_docx, tmp_path / 'out.docx', [{'old': 'Title', 'new': 'Heading'}], True)
