from academic_docx_editor.cli import main, parser
import pytest

from academic_docx_editor.models import MissingDependency
from academic_docx_editor import word_session
from academic_docx_editor.word_session import check_lock


def test_all_commands_have_help(capsys):
    for name in ('doctor', 'audit', 'edit', 'revisions', 'replace-image', 'export-pdf', 'validate'):
        try:
            parser().parse_args([name, '--help'])
        except SystemExit as exc:
            assert exc.code == 0
        else:
            raise AssertionError(name)


def test_lock_file_detected(simple_docx):
    lock = simple_docx.with_name('~$' + simple_docx.name[2:])
    lock.write_text('occupied')
    try:
        check_lock(simple_docx)
    except Exception as exc:
        assert 'lock file' in str(exc)
    else:
        raise AssertionError('lock was not detected')


def test_word_not_installed_is_reported(monkeypatch):
    monkeypatch.setattr(word_session, 'word_executable', lambda: None)
    with pytest.raises(MissingDependency, match='not installed'):
        with word_session.word_application():
            pass
