from cli_anything.python_docx.core.errors import SessionError
from cli_anything.python_docx.core.js_session import JsDocxSession

try:
    from cli_anything.python_docx.core.session import DocxSession
except Exception:  # noqa: BLE001
    # Python fallback session is optional when running JS-only deployments.
    DocxSession = None  # type: ignore[assignment]

__all__ = ["DocxSession", "JsDocxSession", "SessionError"]
