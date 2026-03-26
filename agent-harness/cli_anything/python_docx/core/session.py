"""Core session primitives for the python-docx harness."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from cli_anything.python_docx.utils import python_docx_backend as backend


class SessionError(RuntimeError):
    """Raised when an operation requires unavailable session state."""


class DocxSession:
    def __init__(self) -> None:
        self._document: Any | None = None
        self._path: Path | None = None
        self._undo_stack: list[bytes] = []
        self._redo_stack: list[bytes] = []

    @property
    def path(self) -> Path | None:
        return self._path

    @property
    def has_document(self) -> bool:
        return self._document is not None

    def new_document(self, path: str | Path | None = None, title: str | None = None) -> None:
        self._document = backend.new_document()
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
        self._document = backend.load_document(target)
        self._path = target
        self._undo_stack.clear()
        self._redo_stack.clear()

    def save(self, path: str | Path | None = None) -> Path:
        self._ensure_document()
        target = Path(path) if path is not None else self._path
        if target is None:
            raise SessionError("No target path set. Pass a path to save().")
        target.parent.mkdir(parents=True, exist_ok=True)
        backend.save_document(self._document, target)
        self._path = target
        return target

    def summary(self) -> dict[str, Any]:
        doc = self._ensure_document()
        paragraphs = doc.paragraphs
        tables = doc.tables
        heading_count = 0
        for para in paragraphs:
            style = getattr(para, "style", None)
            style_name = getattr(style, "name", "")
            if isinstance(style_name, str) and style_name.startswith("Heading"):
                heading_count += 1

        return {
            "path": str(self._path) if self._path is not None else None,
            "paragraph_count": len(paragraphs),
            "table_count": len(tables),
            "heading_count": heading_count,
            "undo_depth": len(self._undo_stack),
            "redo_depth": len(self._redo_stack),
        }

    def list_paragraphs(self, limit: int | None = None) -> list[dict[str, Any]]:
        doc = self._ensure_document()
        records: list[dict[str, Any]] = []
        for index, para in enumerate(doc.paragraphs):
            style = getattr(para, "style", None)
            style_name = getattr(style, "name", None)
            records.append(
                {
                    "index": index,
                    "text": para.text,
                    "style": style_name,
                }
            )
        if limit is None:
            return records
        return records[: max(limit, 0)]

    def add_paragraph(self, text: str, style: str | None = None) -> None:
        doc = self._ensure_document()
        self._checkpoint()
        if style:
            doc.add_paragraph(text, style=style)
        else:
            doc.add_paragraph(text)

    def add_heading(self, text: str, level: int = 1) -> None:
        doc = self._ensure_document()
        self._checkpoint()
        doc.add_heading(text, level=level)

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
        line_color: str = "auto",
    ) -> dict[str, int]:
        doc = self._ensure_document()
        self._checkpoint()
        headers = headers or []
        records = records or []

        has_structured_data = bool(headers or records)
        if not has_structured_data:
            if rows is None or cols is None:
                raise SessionError("rows and cols are required when no headers/records are provided.")
            table = doc.add_table(rows=rows, cols=cols)
            if row_lines or column_lines or outer_border:
                self._apply_table_borders(
                    table=table,
                    row_lines=row_lines,
                    column_lines=column_lines,
                    outer_border=outer_border,
                    line_style=line_style,
                    line_size=line_size,
                    line_color=line_color,
                )
            return {"rows": rows, "cols": cols}

        if cols is not None and cols < 1:
            raise SessionError("cols must be >= 1 when provided.")
        if rows is not None and rows < 1:
            raise SessionError("rows must be >= 1 when provided.")

        inferred_cols = cols or 0
        if headers:
            inferred_cols = max(inferred_cols, len(headers))
        for row in records:
            inferred_cols = max(inferred_cols, len(row))

        if inferred_cols < 1:
            raise SessionError("Unable to infer table column count from headers/records.")

        total_rows = (1 if headers else 0) + len(records)
        if rows is not None:
            total_rows = max(total_rows, rows)

        table = doc.add_table(rows=total_rows, cols=inferred_cols)

        row_index = 0
        if headers:
            normalized_header_bg_color = (
                self._normalize_hex_color(header_bg_color, "header_bg_color")
                if header_bg_color
                else None
            )
            for col_index, value in enumerate(headers):
                cell = table.rows[0].cells[col_index]
                self._set_cell_text(cell, value, bold=header_bold)
                if normalized_header_bg_color:
                    self._set_cell_background(cell, normalized_header_bg_color)
            row_index = 1

        for record in records:
            cells = table.rows[row_index].cells
            for col_index, value in enumerate(record):
                cells[col_index].text = value
            row_index += 1

        if row_lines or column_lines or outer_border:
            self._apply_table_borders(
                table=table,
                row_lines=row_lines,
                column_lines=column_lines,
                outer_border=outer_border,
                line_style=line_style,
                line_size=line_size,
                line_color=line_color,
            )

        return {"rows": total_rows, "cols": inferred_cols}

    def _set_cell_text(self, cell: Any, value: str, bold: bool = False) -> None:
        cell.text = value
        if not bold:
            return
        for paragraph in cell.paragraphs:
            if not paragraph.runs:
                paragraph.add_run("")
            for run in paragraph.runs:
                run.bold = True

    def _set_cell_background(self, cell: Any, fill_color: str) -> None:
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = tc_pr.find(qn("w:shd"))
        if shd is None:
            shd = OxmlElement("w:shd")
            tc_pr.append(shd)
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill_color)

    def _normalize_hex_color(self, value: str, field_name: str) -> str:
        normalized = value.strip().lstrip("#").upper()
        if len(normalized) != 6 or any(ch not in "0123456789ABCDEF" for ch in normalized):
            raise SessionError(f"{field_name} must be a 6-digit hex value like 'D9E1F2'.")
        return normalized

    def _apply_table_borders(
        self,
        table: Any,
        row_lines: bool,
        column_lines: bool,
        outer_border: bool,
        line_style: str,
        line_size: int,
        line_color: str,
    ) -> None:
        color = line_color.strip()
        if color.lower() == "auto":
            color = "auto"
        else:
            color = self._normalize_hex_color(color, "line_color")

        tbl = table._tbl
        tbl_pr = tbl.tblPr
        if tbl_pr is None:
            tbl_pr = OxmlElement("w:tblPr")
            tbl.insert(0, tbl_pr)

        tbl_borders = tbl_pr.find(qn("w:tblBorders"))
        if tbl_borders is None:
            tbl_borders = OxmlElement("w:tblBorders")
            tbl_pr.append(tbl_borders)

        border_edges: list[str] = []
        if outer_border:
            border_edges.extend(["top", "left", "bottom", "right"])
        if row_lines:
            border_edges.append("insideH")
        if column_lines:
            border_edges.append("insideV")

        for edge in border_edges:
            edge_elm = tbl_borders.find(qn(f"w:{edge}"))
            if edge_elm is None:
                edge_elm = OxmlElement(f"w:{edge}")
                tbl_borders.append(edge_elm)
            edge_elm.set(qn("w:val"), line_style)
            edge_elm.set(qn("w:sz"), str(line_size))
            edge_elm.set(qn("w:space"), "0")
            edge_elm.set(qn("w:color"), color)

    def set_core_property(self, key: str, value: str) -> None:
        doc = self._ensure_document()
        core = doc.core_properties
        if not hasattr(core, key):
            raise SessionError(f"Unsupported core property: {key}")
        self._checkpoint()
        setattr(core, key, value)

    def undo(self) -> None:
        self._ensure_document()
        if not self._undo_stack:
            raise SessionError("Nothing to undo.")
        current = backend.serialize_document(self._document)
        snapshot = self._undo_stack.pop()
        self._redo_stack.append(current)
        self._document = backend.load_document_from_bytes(snapshot)

    def redo(self) -> None:
        self._ensure_document()
        if not self._redo_stack:
            raise SessionError("Nothing to redo.")
        current = backend.serialize_document(self._document)
        snapshot = self._redo_stack.pop()
        self._undo_stack.append(current)
        self._document = backend.load_document_from_bytes(snapshot)

    def _checkpoint(self) -> None:
        self._ensure_document()
        snapshot = backend.serialize_document(self._document)
        self._undo_stack.append(snapshot)
        self._redo_stack.clear()

    def _ensure_document(self):
        if self._document is None:
            raise SessionError("No active document. Use new/open first.")
        return self._document
