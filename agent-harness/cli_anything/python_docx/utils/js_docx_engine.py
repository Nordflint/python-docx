"""Thin Python wrapper around the Node.js docx engine worker."""

from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path
from typing import Any


def add_bibliography_entry(doc_path: Path, key: str, reference: str, url: str | None = None) -> dict[str, Any]:
    payload = _invoke(
        {
            "command": "addBibliographyEntry",
            "docPath": str(doc_path),
            "key": key,
            "reference": reference,
            "url": url,
        }
    )
    return payload["entry"]


def summary(doc_path: Path) -> dict[str, Any]:
    payload = _invoke(
        {
            "command": "summary",
            "docPath": str(doc_path),
        }
    )
    return payload["summary"]


def list_paragraphs(doc_path: Path, limit: int | None = None) -> list[dict[str, Any]]:
    payload = _invoke(
        {
            "command": "listParagraphs",
            "docPath": str(doc_path),
            "limit": limit,
        }
    )
    return payload["paragraphs"]


def add_paragraph(doc_path: Path, text: str, style: str | None = None) -> dict[str, Any]:
    payload = _invoke(
        {
            "command": "addParagraph",
            "docPath": str(doc_path),
            "text": text,
            "style": style,
        }
    )
    return payload["paragraph"]


def add_heading(doc_path: Path, text: str, level: int = 1) -> dict[str, Any]:
    payload = _invoke(
        {
            "command": "addHeading",
            "docPath": str(doc_path),
            "text": text,
            "level": level,
        }
    )
    return payload["heading"]


def add_table(
    doc_path: Path,
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
    payload = _invoke(
        {
            "command": "addTable",
            "docPath": str(doc_path),
            "rows": rows,
            "cols": cols,
            "headers": headers or [],
            "records": records or [],
            "headerBold": header_bold,
            "headerBgColor": header_bg_color,
            "rowLines": row_lines,
            "columnLines": column_lines,
            "outerBorder": outer_border,
            "lineStyle": line_style,
            "lineSize": line_size,
            "lineColor": line_color,
        }
    )
    return payload["table"]


def set_core_property(doc_path: Path, key: str, value: str) -> dict[str, str]:
    payload = _invoke(
        {
            "command": "setCore",
            "docPath": str(doc_path),
            "key": key,
            "value": value,
        }
    )
    return payload["core"]


def add_frontpage(
    doc_path: Path,
    template: str,
    title: str,
    subtitle: str | None = None,
    author: str | None = None,
    organization: str | None = None,
    date_text: str | None = None,
    include_page_break: bool = True,
    set_core_title: bool = True,
) -> dict[str, Any]:
    payload = _invoke(
        {
            "command": "addFrontpage",
            "docPath": str(doc_path),
            "template": template,
            "title": title,
            "subtitle": subtitle,
            "author": author,
            "organization": organization,
            "dateText": date_text,
            "includePageBreak": include_page_break,
            "setCoreTitle": set_core_title,
        }
    )
    return payload["frontpage"]


def list_bibliography(doc_path: Path) -> list[dict[str, Any]]:
    payload = _invoke(
        {
            "command": "listBibliography",
            "docPath": str(doc_path),
        }
    )
    return payload["entries"]


def add_citation(
    doc_path: Path,
    text: str,
    source_keys: list[str],
    style: str | None = None,
) -> dict[str, Any]:
    payload = _invoke(
        {
            "command": "addCitation",
            "docPath": str(doc_path),
            "text": text,
            "sourceKeys": source_keys,
            "style": style,
        }
    )
    return payload["citation"]


def cite_paragraph(doc_path: Path, paragraph_index: int, source_keys: list[str]) -> dict[str, Any]:
    payload = _invoke(
        {
            "command": "citeParagraph",
            "docPath": str(doc_path),
            "paragraphIndex": paragraph_index,
            "sourceKeys": source_keys,
        }
    )
    return payload["citation"]


def _invoke(request: dict[str, Any]) -> dict[str, Any]:
    node_executable = shutil.which("node")
    if node_executable is None:
        raise RuntimeError("Node.js is required for DOCX_ENGINE=js.")

    engine_script = _engine_script_path()
    if not engine_script.exists():
        raise RuntimeError(f"JS engine script is missing: {engine_script}")

    completed = subprocess.run(
        [node_executable, str(engine_script), json.dumps(request)],
        check=False,
        text=True,
        capture_output=True,
    )
    raw_output = completed.stdout.strip()
    if not raw_output:
        error_text = completed.stderr.strip() or "Unknown JS engine failure."
        raise RuntimeError(error_text)

    try:
        payload = json.loads(raw_output)
    except json.JSONDecodeError as err:
        detail = completed.stderr.strip() or raw_output
        raise RuntimeError(f"Invalid JS engine response: {detail}") from err

    if not payload.get("ok", False):
        raise RuntimeError(payload.get("error", "JS engine command failed."))

    return payload


def _engine_script_path() -> Path:
    return Path(__file__).resolve().parent / "js_engine" / "engine.mjs"
