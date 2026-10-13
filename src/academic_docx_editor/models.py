"""Shared errors and result conventions."""

from __future__ import annotations


class EditorError(Exception):
    exit_code = 2


class UnsafeEdit(EditorError):
    exit_code = 3


class MissingDependency(EditorError):
    exit_code = 4


class ExternalToolError(EditorError):
    exit_code = 5


class ValidationError(EditorError):
    exit_code = 6
