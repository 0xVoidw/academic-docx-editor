"""Own an isolated Word COM process; never attach to a user's Word instance."""

from __future__ import annotations

import contextlib
import os
from pathlib import Path

if os.name == "nt":
    import winreg

from .models import ExternalToolError, MissingDependency, UnsafeEdit


def word_executable() -> str | None:
    if os.name != "nt":
        return None
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE") as key:
                value, _ = winreg.QueryValueEx(key, "")
                if Path(value).is_file():
                    return value
        except OSError:
            pass
    return None


def check_lock(path: Path) -> None:
    path = Path(path)
    owner = path.with_name("~$" + path.name[2:])
    if owner.exists():
        raise UnsafeEdit(f"Word lock file exists: {owner}")


@contextlib.contextmanager
def word_application():
    if os.name != "nt" or not word_executable():
        raise MissingDependency("Microsoft Word is not installed or registered")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise MissingDependency("pywin32 is required: pip install 'academic-docx-editor[word]'") from exc
    initialized = False
    app = None
    try:
        pythoncom.CoInitialize()
        initialized = True
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        yield app
    except Exception as exc:
        if isinstance(exc, (MissingDependency, UnsafeEdit, ExternalToolError)):
            raise
        raise ExternalToolError(f"Word COM failure: {exc}") from exc
    finally:
        if app is not None:
            try:
                app.Quit(SaveChanges=0)
            except Exception:
                pass
        if initialized:
            pythoncom.CoUninitialize()


@contextlib.contextmanager
def opened_document(app, path: Path, read_only: bool):
    document = None
    try:
        document = app.Documents.Open(FileName=str(Path(path).resolve()), ReadOnly=read_only,
                                      AddToRecentFiles=False, ConfirmConversions=False, Visible=False)
        yield document
    finally:
        if document is not None:
            document.Close(SaveChanges=0)
