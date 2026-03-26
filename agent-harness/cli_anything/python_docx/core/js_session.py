"""JS-engine-backed session primitives for the python-docx harness."""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any, Callable, TypeVar

from cli_anything.python_docx.core.session import SessionError
from cli_anything.python_docx.utils import js_docx_engine, python_docx_backend as backend

T = TypeVar("T")


class JsDocxSession:
    FRONTPAGE_TEMPLATES = ("clean", "corporate", "academic")

    def __init__(self) -> None:
        self._document_bytes: bytes | None = None
        self._path: Path | None = None
        self._undo_stack: list[bytes] = []
        self._redo_stack: list[bytes] = []

    @property
    def path(self) -> Path | None:
        return self._path

    @property
    def has_document(self) -> bool:
        return self._document_bytes is not None

    def new_document(self, path: str | Path | None = None, title: str | None = None) -> None:
        document = backend.new_document()
        self._document_bytes = backend.serialize_document(document)
        self._path = Path(path) if path is not None else None
        self._undo_stack.clear()
        self._redo_stack.clear()
        if title:
            self.set_core_property("title", title)
        if path is not None:
            self.save(path)

    def open_document(self, path: str | Path) -> None:
        target = Path(path)
        if not target.exists():
            raise SessionError(f"Document does not exist: {target}")
        self._document_bytes = target.read_bytes()
        self._path = target
        self._undo_stack.clear()
        self._redo_stack.clear()

    def save(self, path: str | Path | None = None) -> Path:
        document_bytes = self._ensure_document()
        target = Path(path) if path is not None else self._path
        if target is None:
            raise SessionError("No target path set. Pass a path to save().")
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(document_bytes)
        self._path = target
        return target

    def summary(self) -> dict[str, Any]:
        summary = self._read_with_temp_doc(lambda doc_path: js_docx_engine.summary(doc_path))
        summary["path"] = str(self._path) if self._path is not None else None
        summary["undo_depth"] = len(self._undo_stack)
        summary["redo_depth"] = len(self._redo_stack)
        return summary

    def list_paragraphs(self, limit: int | None = None) -> list[dict[str, Any]]:
        return self._read_with_temp_doc(lambda doc_path: js_docx_engine.list_paragraphs(doc_path, limit=limit))

    def add_paragraph(self, text: str, style: str | None = None) -> None:
        self._mutate_with_temp_doc(lambda doc_path: js_docx_engine.add_paragraph(doc_path, text=text, style=style))

    def add_heading(self, text: str, level: int = 1) -> None:
        self._mutate_with_temp_doc(lambda doc_path: js_docx_engine.add_heading(doc_path, text=text, level=level))

    def add_table(
        self,
        rows: int | None = None,
        cols: int | None = None,
        headers: list[str] | None = None,
        records: list[list[str]] | None = None,
        header_bold: bool = False,
        header_bg_color: str | None = None,
        row_lines: bool = False,
        column_lines: bool = False,
        outer_border: bool = False,
        line_style: str = "single",
        line_size: int = 8,
        line_color: str = "main",
    ) -> dict[str, int]:
        return self._mutate_with_temp_doc(
            lambda doc_path: js_docx_engine.add_table(
                doc_path=doc_path,
                rows=rows,
                cols=cols,
                headers=headers,
                records=records,
                header_bold=header_bold,
                header_bg_color=header_bg_color,
                row_lines=row_lines,
                column_lines=column_lines,
                outer_border=outer_border,
                line_style=line_style,
                line_size=line_size,
                line_color=line_color,
            )
        )

    def list_frontpage_templates(self) -> list[str]:
        return list(self.FRONTPAGE_TEMPLATES)

    def add_frontpage(
        self,
        template: str,
        title: str,
        subtitle: str | None = None,
        author: str | None = None,
        organization: str | None = None,
        date_text: str | None = None,
        include_page_break: bool = True,
        set_core_title: bool = True,
    ) -> dict[str, Any]:
        return self._mutate_with_temp_doc(
            lambda doc_path: js_docx_engine.add_frontpage(
                doc_path=doc_path,
                template=template,
                title=title,
                subtitle=subtitle,
                author=author,
                organization=organization,
                date_text=date_text,
                include_page_break=include_page_break,
                set_core_title=set_core_title,
            )
        )

    def add_bibliography_entry(
        self,
        key: str,
        reference: str,
        url: str | None = None,
    ) -> dict[str, Any]:
        return self._mutate_with_temp_doc(
            lambda doc_path: js_docx_engine.add_bibliography_entry(
                doc_path=doc_path,
                key=key,
                reference=reference,
                url=url,
            )
        )

    def list_bibliography(self) -> list[dict[str, Any]]:
        return self._read_with_temp_doc(js_docx_engine.list_bibliography)

    def add_citation(self, text: str, source_keys: list[str], style: str | None = None) -> dict[str, Any]:
        return self._mutate_with_temp_doc(
            lambda doc_path: js_docx_engine.add_citation(
                doc_path=doc_path,
                text=text,
                source_keys=source_keys,
                style=style,
            )
        )

    def cite_paragraph(self, paragraph_index: int, source_keys: list[str]) -> dict[str, Any]:
        return self._mutate_with_temp_doc(
            lambda doc_path: js_docx_engine.cite_paragraph(
                doc_path=doc_path,
                paragraph_index=paragraph_index,
                source_keys=source_keys,
            )
        )

    def set_core_property(self, key: str, value: str) -> None:
        self._mutate_with_temp_doc(lambda doc_path: js_docx_engine.set_core_property(doc_path, key=key, value=value))

    def undo(self) -> None:
        self._ensure_document()
        if not self._undo_stack:
            raise SessionError("Nothing to undo.")
        current = self._ensure_document()
        snapshot = self._undo_stack.pop()
        self._redo_stack.append(current)
        self._document_bytes = snapshot

    def redo(self) -> None:
        self._ensure_document()
        if not self._redo_stack:
            raise SessionError("Nothing to redo.")
        current = self._ensure_document()
        snapshot = self._redo_stack.pop()
        self._undo_stack.append(current)
        self._document_bytes = snapshot

    def _checkpoint(self) -> None:
        snapshot = self._ensure_document()
        self._undo_stack.append(snapshot)
        self._redo_stack.clear()

    def _ensure_document(self) -> bytes:
        if self._document_bytes is None:
            raise SessionError("No active document. Use new/open first.")
        return self._document_bytes

    def _read_with_temp_doc(self, operation: Callable[[Path], T]) -> T:
        document_bytes = self._ensure_document()
        with tempfile.TemporaryDirectory(prefix="js-docx-session-") as temp_dir:
            doc_path = Path(temp_dir) / "active.docx"
            doc_path.write_bytes(document_bytes)
            return operation(doc_path)

    def _mutate_with_temp_doc(self, operation: Callable[[Path], T]) -> T:
        self._checkpoint()
        document_bytes = self._ensure_document()
        with tempfile.TemporaryDirectory(prefix="js-docx-session-") as temp_dir:
            doc_path = Path(temp_dir) / "active.docx"
            doc_path.write_bytes(document_bytes)
            result = operation(doc_path)
            self._document_bytes = doc_path.read_bytes()
            return result
