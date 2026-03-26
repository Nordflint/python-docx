"""Shared exceptions for python-docx harness sessions."""

from __future__ import annotations


class SessionError(RuntimeError):
    """Raised when an operation requires unavailable session state."""
