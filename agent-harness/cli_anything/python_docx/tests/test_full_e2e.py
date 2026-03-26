from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn


def _cli_cmd() -> list[str]:
    exe = shutil.which("cli-anything-python-docx")
    if exe:
        return [exe]
    return [sys.executable, "-m", "cli_anything.python_docx"]


def _run(
    args: list[str],
    input_text: str | None = None,
    extra_env: dict[str, str] | None = None,
) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    if extra_env:
        env.update(extra_env)
    return subprocess.run(
        _cli_cmd() + args,
        check=True,
        capture_output=True,
        text=True,
        input=input_text,
        env=env,
    )


def _json_lines(output: str) -> list[dict]:
    rows = []
    for line in output.splitlines():
        line = line.strip()
        start = line.find("{")
        if start < 0:
            continue
        candidate = line[start:]
        if not candidate.endswith("}"):
            continue
        rows.append(json.loads(candidate))
    return rows


def _js_engine_ready() -> bool:
    if shutil.which("node") is None:
        return False
    harness_root = Path(__file__).resolve().parents[3]
    return (harness_root / "node_modules").is_dir()


def it_runs_one_shot_json_workflow(tmp_path: Path) -> None:
    doc_path = tmp_path / "workflow.docx"

    created = _run(["--json", "new", str(doc_path), "--title", "Workflow"])
    created_rows = _json_lines(created.stdout)
    assert created_rows and created_rows[-1]["ok"] is True

    _run(["add-paragraph", "Hello from e2e", "--doc", str(doc_path)])
    _run(["add-heading", "Milestone", "--level", "2", "--doc", str(doc_path)])

    summary = _run(["--json", "summary", "--doc", str(doc_path)])
    summary_rows = _json_lines(summary.stdout)
    assert summary_rows

    payload = summary_rows[-1]
    assert payload["summary"]["paragraph_count"] == 2
    assert payload["summary"]["heading_count"] == 1


def it_reports_effective_engine() -> None:
    completed = _run(["--json", "engine"])
    rows = _json_lines(completed.stdout)
    assert rows
    payload = rows[-1]
    assert payload["action"] == "engine"
    assert payload["engine"]["active"] in {"python", "js"}
    assert isinstance(payload["engine"]["js_runtime_ready"], bool)
    assert isinstance(payload["engine"]["python_runtime_ready"], bool)


def it_runs_default_repl_mode_with_undo() -> None:
    repl_script = "\n".join(
        [
            "new",
            "add-paragraph \"Temp line\"",
            "undo",
            "summary",
            "exit",
            "",
        ]
    )

    completed = _run(["--json"], input_text=repl_script)
    rows = _json_lines(completed.stdout)

    summary_rows = [row for row in rows if row.get("action") == "summary"]
    assert summary_rows
    assert summary_rows[-1]["summary"]["paragraph_count"] == 0


def it_adds_table_records_from_cli(tmp_path: Path) -> None:
    doc_path = tmp_path / "table-records.docx"
    _run(["new", str(doc_path)])
    _run(
        [
            "add-table",
            "--doc",
            str(doc_path),
            "--header",
            "Qty",
            "--header",
            "Id",
            "--header",
            "Desc",
            "--header-bold",
            "--record",
            "3|101|Spam",
            "--record",
            "7|422|Eggs",
            "--record",
            "4|631|Spam, spam, eggs, and spam",
        ]
    )

    doc = Document(str(doc_path))
    table = doc.tables[0]
    assert len(table.rows) == 4
    assert len(table.columns) == 3
    assert table.cell(0, 0).text == "Qty"
    assert table.cell(1, 0).text == "3"
    assert table.cell(3, 2).text == "Spam, spam, eggs, and spam"
    assert table.cell(0, 0).paragraphs[0].runs[0].bold is True
    shd = table.cell(0, 0)._tc.get_or_add_tcPr().find(qn("w:shd"))
    assert shd is not None
    assert shd.get(qn("w:fill")) == "05206E"
    assert str(table.cell(0, 0).paragraphs[0].runs[0].font.color.rgb) == "FFFFFF"


def it_adds_row_and_column_lines_from_cli(tmp_path: Path) -> None:
    doc_path = tmp_path / "table-lines.docx"
    _run(["new", str(doc_path)])
    _run(
        [
            "add-table",
            "--doc",
            str(doc_path),
            "--rows",
            "2",
            "--cols",
            "2",
            "--row-lines",
            "--column-lines",
            "--outer-border",
            "--line-style",
            "single",
            "--line-size",
            "8",
            "--line-color",
            "main",
        ]
    )

    doc = Document(str(doc_path))
    table = doc.tables[0]
    tbl_borders = table._tbl.tblPr.find(qn("w:tblBorders"))
    assert tbl_borders is not None
    assert tbl_borders.find(qn("w:insideH")) is not None
    assert tbl_borders.find(qn("w:insideV")) is not None
    assert tbl_borders.find(qn("w:insideH")).get(qn("w:color")) == "05206E"


def it_inserts_frontpage_template_from_cli(tmp_path: Path) -> None:
    doc_path = tmp_path / "frontpage-cli.docx"
    _run(["new", str(doc_path)])
    _run(["add-paragraph", "--doc", str(doc_path), "Existing body content"])
    _run(
        [
            "add-frontpage",
            "--doc",
            str(doc_path),
            "--template",
            "clean",
            "--title",
            "Operations Review",
            "--subtitle",
            "Q1 2026",
            "--author",
            "CLI Agent",
            "--organization",
            "Nordflint",
            "--date-text",
            "2026-03-26",
        ]
    )

    doc = Document(str(doc_path))
    assert doc.paragraphs[0].text == "Operations Review"
    assert str(doc.paragraphs[0].runs[0].font.color.rgb) == "05206E"
    body_index = next(i for i, para in enumerate(doc.paragraphs) if para.text == "Existing body content")
    assert body_index > 0
    assert doc.core_properties.title == "Operations Review"


def it_uses_blue_background_and_white_font_for_corporate_frontpage(tmp_path: Path) -> None:
    doc_path = tmp_path / "frontpage-corporate.docx"
    _run(["new", str(doc_path)])
    _run(
        [
            "add-frontpage",
            "--doc",
            str(doc_path),
            "--template",
            "corporate",
            "--title",
            "Corporate Report",
            "--subtitle",
            "Q2 2026",
            "--author",
            "CLI Agent",
            "--organization",
            "Nordflint",
            "--date-text",
            "2026-03-26",
        ]
    )

    doc = Document(str(doc_path))
    assert doc.paragraphs[0].text == "NORDFLINT"
    assert str(doc.paragraphs[0].runs[0].font.color.rgb) == "FFFFFF"
    para0_shd = doc.paragraphs[0]._p.get_or_add_pPr().find(qn("w:shd"))
    assert para0_shd is not None
    assert para0_shd.get(qn("w:fill")) == "05206E"


def it_adds_bibliography_and_citations_from_cli(tmp_path: Path) -> None:
    doc_path = tmp_path / "citations-cli.docx"
    py_env = {"DOCX_ENGINE": "python"}
    _run(["new", str(doc_path)])
    _run(["add-paragraph", "--doc", str(doc_path), "The regional labor market improved year over year."], extra_env=py_env)
    _run(
        [
            "add-bibliography-entry",
            "--doc",
            str(doc_path),
            "--url",
            "https://example.com/statsdk2025",
            "statsdk2025",
            "Statistics Denmark labour bulletin (2025)",
        ],
        extra_env=py_env,
    )
    _run(
        [
            "add-bibliography-entry",
            "--doc",
            str(doc_path),
            "oecd2024",
            "OECD Employment Outlook (2024)",
        ],
        extra_env=py_env,
    )
    _run(
        [
            "cite-paragraph",
            "--doc",
            str(doc_path),
            "--index",
            "0",
            "--source-key",
            "statsdk2025",
            "--source-key",
            "oecd2024",
        ],
        extra_env=py_env,
    )
    _run(
        [
            "add-citation",
            "--doc",
            str(doc_path),
            "--source-key",
            "oecd2024",
            "OECD data also shows lower youth unemployment.",
        ],
        extra_env=py_env,
    )

    listed = _run(["--json", "list-bibliography", "--doc", str(doc_path)], extra_env=py_env)
    listed_rows = _json_lines(listed.stdout)
    assert listed_rows
    entries = listed_rows[-1]["entries"]
    assert len(entries) == 2
    assert entries[0]["key"] == "statsdk2025"
    assert entries[1]["key"] == "oecd2024"

    doc = Document(str(doc_path))
    texts = [paragraph.text for paragraph in doc.paragraphs]
    assert texts[0].endswith("[1, 2]")
    bibliography_index = texts.index("Bibliography")
    assert texts[bibliography_index + 1].startswith("[1] statsdk2025:")
    assert "(URL: https://example.com/statsdk2025)" in texts[bibliography_index + 1]
    assert texts[bibliography_index + 2].startswith("[2] oecd2024:")
    assert any(text.startswith("OECD data also shows") and text.endswith("[2]") for text in texts)


@pytest.mark.skipif(
    not _js_engine_ready(),
    reason="JS engine dependencies are not installed. Run `npm install` in agent-harness/ first.",
)
def it_prefers_js_engine_in_auto_mode_when_available(tmp_path: Path) -> None:
    doc_path = tmp_path / "auto-engine-prefers-js.docx"

    _run(["new", str(doc_path)])
    _run(["add-bibliography-entry", "--doc", str(doc_path), "src1", "Source One"])
    _run(["add-citation", "--doc", str(doc_path), "--source-key", "src1", "Auto mode claim text."])

    with zipfile.ZipFile(doc_path) as archive:
        assert "word/footnotes.xml" in archive.namelist()
        document_xml = archive.read("word/document.xml").decode("utf-8")
    assert re.search(r"<w:footnoteReference[^>]*w:id=\"1\"", document_xml)


@pytest.mark.skipif(
    not _js_engine_ready(),
    reason="JS engine dependencies are not installed. Run `npm install` in agent-harness/ first.",
)
def it_runs_js_engine_repl_session_with_undo_and_save(tmp_path: Path) -> None:
    doc_path = tmp_path / "js-repl-session.docx"
    js_env = {"DOCX_ENGINE": "js"}
    repl_script = "\n".join(
        [
            "new",
            "add-paragraph \"Session line\"",
            "undo",
            f"save \"{doc_path}\"",
            "summary",
            "exit",
            "",
        ]
    )

    completed = _run(["--json"], input_text=repl_script, extra_env=js_env)
    rows = _json_lines(completed.stdout)
    summary_rows = [row for row in rows if row.get("action") == "summary"]
    assert summary_rows
    assert summary_rows[-1]["summary"]["paragraph_count"] == 0
    assert doc_path.exists()

    doc = Document(str(doc_path))
    assert len(doc.paragraphs) == 0


@pytest.mark.skipif(
    not _js_engine_ready(),
    reason="JS engine dependencies are not installed. Run `npm install` in agent-harness/ first.",
)
def it_runs_summary_and_paragraph_commands_with_js_engine(tmp_path: Path) -> None:
    doc_path = tmp_path / "js-engine-summary.docx"
    js_env = {"DOCX_ENGINE": "js"}

    _run(["new", str(doc_path)])
    _run(["add-paragraph", "--doc", str(doc_path), "Alpha line"], extra_env=js_env)
    _run(["add-heading", "--doc", str(doc_path), "--level", "2", "Milestone"], extra_env=js_env)

    summary = _run(["--json", "summary", "--doc", str(doc_path)], extra_env=js_env)
    summary_rows = _json_lines(summary.stdout)
    assert summary_rows
    summary_payload = summary_rows[-1]["summary"]
    assert summary_payload["paragraph_count"] == 2
    assert summary_payload["heading_count"] == 1

    listed = _run(
        ["--json", "list-paragraphs", "--doc", str(doc_path)],
        extra_env=js_env,
    )
    listed_rows = _json_lines(listed.stdout)
    assert listed_rows
    paragraphs = listed_rows[-1]["paragraphs"]
    assert len(paragraphs) == 2
    assert paragraphs[0]["text"] == "Alpha line"
    assert paragraphs[1]["text"] == "Milestone"

    doc = Document(str(doc_path))
    assert doc.paragraphs[0].text == "Alpha line"
    assert doc.paragraphs[1].text == "Milestone"
    assert str(doc.paragraphs[1].runs[0].font.color.rgb) == "05206E"


@pytest.mark.skipif(
    not _js_engine_ready(),
    reason="JS engine dependencies are not installed. Run `npm install` in agent-harness/ first.",
)
def it_adds_tables_with_js_engine(tmp_path: Path) -> None:
    doc_path = tmp_path / "js-engine-table.docx"
    js_env = {"DOCX_ENGINE": "js"}

    _run(["new", str(doc_path)])
    _run(
        [
            "add-table",
            "--doc",
            str(doc_path),
            "--header",
            "Qty",
            "--header",
            "Id",
            "--header",
            "Desc",
            "--header-bold",
            "--record",
            "3|101|Spam",
            "--record",
            "7|422|Eggs",
            "--row-lines",
            "--column-lines",
            "--outer-border",
            "--line-style",
            "single",
            "--line-size",
            "8",
            "--line-color",
            "main",
        ],
        extra_env=js_env,
    )

    doc = Document(str(doc_path))
    table = doc.tables[0]
    assert len(table.rows) == 3
    assert len(table.columns) == 3
    assert table.cell(0, 0).text == "Qty"
    assert table.cell(1, 0).text == "3"
    assert table.cell(2, 2).text == "Eggs"
    assert table.cell(0, 0).paragraphs[0].runs[0].bold is True
    shd = table.cell(0, 0)._tc.get_or_add_tcPr().find(qn("w:shd"))
    assert shd is not None
    assert shd.get(qn("w:fill")) == "05206E"
    assert str(table.cell(0, 0).paragraphs[0].runs[0].font.color.rgb) == "FFFFFF"
    tbl_borders = table._tbl.tblPr.find(qn("w:tblBorders"))
    assert tbl_borders is not None
    assert tbl_borders.find(qn("w:insideH")) is not None
    assert tbl_borders.find(qn("w:insideV")) is not None
    assert tbl_borders.find(qn("w:insideH")).get(qn("w:color")) == "05206E"


@pytest.mark.skipif(
    not _js_engine_ready(),
    reason="JS engine dependencies are not installed. Run `npm install` in agent-harness/ first.",
)
def it_sets_core_properties_with_js_engine(tmp_path: Path) -> None:
    doc_path = tmp_path / "js-engine-core.docx"
    js_env = {"DOCX_ENGINE": "js"}

    _run(["new", str(doc_path)])
    _run(["set-core", "--doc", str(doc_path), "title", "JS Engine Report"], extra_env=js_env)
    _run(["set-core", "--doc", str(doc_path), "author", "CLI Agent"], extra_env=js_env)
    _run(["set-core", "--doc", str(doc_path), "keywords", "economy, forecast"], extra_env=js_env)

    doc = Document(str(doc_path))
    assert doc.core_properties.title == "JS Engine Report"
    assert doc.core_properties.author == "CLI Agent"
    assert doc.core_properties.keywords == "economy, forecast"


@pytest.mark.skipif(
    not _js_engine_ready(),
    reason="JS engine dependencies are not installed. Run `npm install` in agent-harness/ first.",
)
def it_inserts_frontpage_with_js_engine(tmp_path: Path) -> None:
    doc_path = tmp_path / "js-engine-frontpage.docx"
    js_env = {"DOCX_ENGINE": "js"}

    _run(["new", str(doc_path)])
    _run(["add-paragraph", "--doc", str(doc_path), "Existing body content"])
    _run(
        [
            "add-frontpage",
            "--doc",
            str(doc_path),
            "--template",
            "corporate",
            "--title",
            "Operations Review",
            "--subtitle",
            "Q1 2026",
            "--author",
            "CLI Agent",
            "--organization",
            "Nordflint",
            "--date-text",
            "2026-03-26",
        ],
        extra_env=js_env,
    )

    doc = Document(str(doc_path))
    assert doc.paragraphs[0].text == "NORDFLINT"
    assert doc.paragraphs[1].text == "Operations Review"
    assert str(doc.paragraphs[1].runs[0].font.color.rgb) == "FFFFFF"
    body_index = next(i for i, para in enumerate(doc.paragraphs) if para.text == "Existing body content")
    assert body_index > 1
    assert doc.core_properties.title == "Operations Review"


@pytest.mark.skipif(
    not _js_engine_ready(),
    reason="JS engine dependencies are not installed. Run `npm install` in agent-harness/ first.",
)
def it_writes_footnote_citations_with_js_engine(tmp_path: Path) -> None:
    doc_path = tmp_path / "citations-js-engine.docx"
    js_env = {"DOCX_ENGINE": "js"}

    _run(["new", str(doc_path)])
    _run(["add-paragraph", "--doc", str(doc_path), "Primary claim text."])
    _run(
        [
            "add-bibliography-entry",
            "--doc",
            str(doc_path),
            "--url",
            "https://example.com/agency2025",
            "agency2025",
            "Agency Annual Report 2025",
        ],
        extra_env=js_env,
    )
    _run(
        [
            "cite-paragraph",
            "--doc",
            str(doc_path),
            "--index",
            "0",
            "--source-key",
            "agency2025",
        ],
        extra_env=js_env,
    )
    _run(
        [
            "add-citation",
            "--doc",
            str(doc_path),
            "--source-key",
            "agency2025",
            "Follow-up claim text.",
        ],
        extra_env=js_env,
    )

    listed = _run(
        ["--json", "list-bibliography", "--doc", str(doc_path)],
        extra_env=js_env,
    )
    listed_rows = _json_lines(listed.stdout)
    assert listed_rows
    assert listed_rows[-1]["entries"][0]["key"] == "agency2025"

    paragraphs = _run(
        ["--json", "list-paragraphs", "--doc", str(doc_path)],
        extra_env=js_env,
    )
    paragraph_rows = _json_lines(paragraphs.stdout)
    assert paragraph_rows
    rendered = [row["text"] for row in paragraph_rows[-1]["paragraphs"]]
    assert any(text.endswith("[1]") for text in rendered)

    with zipfile.ZipFile(doc_path) as archive:
        assert "word/footnotes.xml" in archive.namelist()
        document_xml = archive.read("word/document.xml").decode("utf-8")
        footnotes_xml = archive.read("word/footnotes.xml").decode("utf-8")

    assert re.search(r"<w:footnoteReference[^>]*w:id=\"1\"", document_xml)
    assert "Agency Annual Report 2025" in footnotes_xml
