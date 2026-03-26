from cli_anything.python_docx.utils import js_docx_engine

try:
    from cli_anything.python_docx.utils import python_docx_backend
except Exception:  # noqa: BLE001
    # Python backend is optional for JS-first deployments.
    python_docx_backend = None  # type: ignore[assignment]

__all__ = ["python_docx_backend", "js_docx_engine"]
