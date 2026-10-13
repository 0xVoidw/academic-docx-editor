from academic_docx_editor.audit import inspect_docx, markdown_report
from academic_docx_editor.models import ValidationError
from academic_docx_editor.ooxml import Package


def test_audit_rich_structure(rich_docx):
    report = inspect_docx(rich_docx)
    assert report['counts']['images'] == 2
    assert report['counts']['tables'] == 1
    assert report['counts']['formulas'] == 1
    assert report['counts']['hyperlinks'] == 1
    assert report['images'][0]['kind'] == 'inline'
    assert report['images'][1]['kind'] == 'floating'
    assert report['cross_references'][0]['code'].startswith('REF')
    assert '论文页眉' in report['headers_footers']['word/header1.xml']
    assert report['properties']['docProps/core.xml']['title'] == '独立测试论文'
    assert any(p['style_id'] == 'Heading2' for p in report['paragraphs'])
    assert '研究标题' in markdown_report(report)


def test_corrupt_docx_rejected(tmp_path):
    path = tmp_path / 'bad.docx'
    path.write_bytes(b'not a zip')
    try:
        Package(path)
    except ValidationError:
        pass
    else:
        raise AssertionError('corrupt package was accepted')
