from __future__ import annotations

import csv
import json
import os
import shlex
import shutil
from functools import lru_cache
from pathlib import Path
from typing import Any, cast

import click

from cli_anything.python_docx.core import JsDocxSession, SessionError


def _format_summary(summary: dict[str, Any]) -> str:
    lines = [
        f"path: {summary['path']}",
        f"paragraph_count: {summary['paragraph_count']}",
        f"table_count: {summary['table_count']}",
        f"heading_count: {summary['heading_count']}",
        f"undo_depth: {summary['undo_depth']}",
        f"redo_depth: {summary['redo_depth']}",
    ]
    return "\n".join(lines)


def _emit(ctx: click.Context, payload: dict[str, Any], text: str | None = None) -> None:
    if ctx.obj.get("json_output", False):
        click.echo(json.dumps(payload, sort_keys=True))
        return
    if text is not None:
        click.echo(text)
        return
    click.echo(json.dumps(payload, indent=2, sort_keys=True))


def _session_from_context(ctx: click.Context) -> Any:
    return ctx.obj["session"]


def _open_if_requested(session: Any, doc: str | None) -> bool:
    if not doc:
        return False
    session.open_document(Path(doc))
    return True


def _require_session_doc(session: Any) -> None:
    if not session.has_document:
        raise click.UsageError("No active document. Use new/open or pass --doc.")


def _parse_records(record_values: tuple[str, ...], delimiter: str) -> list[list[str]]:
    rows: list[list[str]] = []
    for value in record_values:
        parsed = next(csv.reader([value], delimiter=delimiter))
        rows.append(parsed)
    return rows


def _require_source_keys(source_keys: tuple[str, ...]) -> list[str]:
    if not source_keys:
        raise click.UsageError("Provide at least one --source-key.")
    return list(source_keys)


def _active_engine() -> str:
    configured = str(os.environ.get("DOCX_ENGINE", "auto")).strip().lower()
    if configured == "auto":
        return "js" if _js_runtime_ready() else "python"
    if configured in {"js", "python"}:
        return configured
    return "js" if _js_runtime_ready() else "python"


def _configured_engine() -> str:
    return str(os.environ.get("DOCX_ENGINE", "auto")).strip().lower()


def _js_runtime_ready() -> bool:
    if shutil.which("node") is None:
        return False
    engine_script = Path(__file__).resolve().parent / "utils" / "js_engine" / "engine.mjs"
    if not engine_script.exists():
        return False
    node_modules_dir = Path(__file__).resolve().parents[2] / "node_modules"
    return node_modules_dir.exists()


@lru_cache(maxsize=1)
def _python_runtime_ready() -> bool:
    try:
        from cli_anything.python_docx.core.session import DocxSession  # noqa: F401
    except Exception:  # noqa: BLE001
        return False
    return True


def _run_repl(root: click.Command, ctx_obj: dict[str, Any]) -> None:
    click.echo("python-docx REPL. Type 'help' for commands, 'exit' to quit.")
    while True:
        try:
            line = input("python-docx> ").strip()
        except EOFError:
            click.echo()
            break
        except KeyboardInterrupt:
            click.echo()
            continue

        if not line:
            continue

        if line in {"exit", "quit"}:
            break

        if line == "help":
            click.echo(
                "Commands: new, open, save, summary, list-paragraphs, add-paragraph, "
                "add-heading, add-table, frontpage-templates, add-frontpage, "
                "add-bibliography-entry, list-bibliography, add-citation, cite-paragraph, "
                "set-core, undo, redo, engine, repl"
            )
            click.echo("Use 'json on' or 'json off' to toggle JSON output.")
            continue

        if line.startswith("json "):
            mode = line.split(maxsplit=1)[1].strip().lower()
            if mode not in {"on", "off"}:
                click.echo("json mode must be 'on' or 'off'")
                continue
            ctx_obj["json_output"] = mode == "on"
            click.echo(f"json output {'enabled' if ctx_obj['json_output'] else 'disabled'}")
            continue

        try:
            tokens = shlex.split(line)
        except ValueError as err:
            click.echo(f"Parse error: {err}")
            continue

        if not tokens:
            continue

        args = (["--json"] if ctx_obj.get("json_output", False) else []) + tokens
        try:
            root.main(
                args=args,
                prog_name="cli-anything-python-docx",
                obj=ctx_obj,
                standalone_mode=False,
            )
        except click.ClickException as err:
            err.show()
        except SessionError as err:
            click.echo(f"Error: {err}")
        except Exception as err:  # noqa: BLE001
            click.echo(f"Error: {err}")


def _build_session() -> Any:
    active_engine = _active_engine()
    if active_engine == "js":
        if not _js_runtime_ready():
            raise click.ClickException(
                "JS engine is unavailable. Install Node dependencies in agent-harness/ "
                "or set DOCX_ENGINE=python."
            )
        return JsDocxSession()
    if not _python_runtime_ready():
        if _js_runtime_ready():
            raise click.ClickException(
                "Python engine is unavailable. Install python-docx or set DOCX_ENGINE=js."
            )
        raise click.ClickException(
            "No DOCX engine runtime is available. Install Node dependencies for JS mode "
            "or install python-docx for Python mode."
        )
    try:
        from cli_anything.python_docx.core.session import DocxSession
    except Exception as err:  # noqa: BLE001
        raise click.ClickException(
            "Python engine is unavailable. Install python-docx or set DOCX_ENGINE=js."
        ) from err
    return cast(Any, DocxSession())


@click.group(invoke_without_command=True)
@click.option("--json", "json_output", is_flag=True, help="Emit JSON output.")
@click.pass_context
def cli(ctx: click.Context, json_output: bool) -> None:
    """CLI-Anything harness for python-docx."""
    if ctx.obj is None:
        ctx.obj = {}
    if "session" not in ctx.obj:
        ctx.obj["session"] = _build_session()
    if json_output:
        ctx.obj["json_output"] = True
    else:
        ctx.obj.setdefault("json_output", False)

    if ctx.invoked_subcommand is None:
        _run_repl(cli, ctx.obj)


@cli.command("repl")
@click.pass_context
def repl(ctx: click.Context) -> None:
    """Start interactive mode."""
    _run_repl(cli, ctx.obj)


@cli.command("engine")
@click.pass_context
def engine_command(ctx: click.Context) -> None:
    """Show effective engine selection details."""
    payload = {
        "ok": True,
        "action": "engine",
        "engine": {
            "configured": _configured_engine(),
            "active": _active_engine(),
            "js_runtime_ready": _js_runtime_ready(),
            "python_runtime_ready": _python_runtime_ready(),
        },
    }
    _emit(
        ctx,
        payload,
        text=(
            f"configured={payload['engine']['configured']} "
            f"active={payload['engine']['active']} "
            f"js_runtime_ready={payload['engine']['js_runtime_ready']} "
            f"python_runtime_ready={payload['engine']['python_runtime_ready']}"
        ),
    )


@cli.command("new")
@click.argument("path", required=False, type=click.Path(path_type=Path))
@click.option("--title", default=None, help="Optional core title metadata.")
@click.pass_context
def new_command(ctx: click.Context, path: Path | None, title: str | None) -> None:
    """Create a new document and optionally save it."""
    session = _session_from_context(ctx)
    session.new_document(path=path, title=title)
    payload = {
        "ok": True,
        "action": "new",
        "path": str(session.path) if session.path else None,
        "title": title,
    }
    _emit(ctx, payload, text=f"Created document at {session.path}" if session.path else "Created new in-memory document")


@cli.command("open")
@click.argument("path", type=click.Path(exists=True, path_type=Path))
@click.pass_context
def open_command(ctx: click.Context, path: Path) -> None:
    """Open an existing .docx document."""
    session = _session_from_context(ctx)
    session.open_document(path)
    summary = session.summary()
    payload = {"ok": True, "action": "open", "summary": summary}
    _emit(ctx, payload, text=f"Opened {path}\n{_format_summary(summary)}")


@cli.command("save")
@click.argument("path", required=False, type=click.Path(path_type=Path))
@click.pass_context
def save_command(ctx: click.Context, path: Path | None) -> None:
    """Save the active document."""
    session = _session_from_context(ctx)
    saved = session.save(path)
    payload = {"ok": True, "action": "save", "path": str(saved)}
    _emit(ctx, payload, text=f"Saved document: {saved}")


@cli.command("summary")
@click.option("--doc", type=click.Path(exists=True, path_type=Path), default=None, help="Open this document before reporting summary.")
@click.pass_context
def summary_command(ctx: click.Context, doc: Path | None) -> None:
    """Show document summary."""
    session = _session_from_context(ctx)
    _open_if_requested(session, str(doc) if doc else None)
    _require_session_doc(session)
    summary = session.summary()
    payload = {"ok": True, "action": "summary", "summary": summary}
    _emit(ctx, payload, text=_format_summary(summary))


@cli.command("list-paragraphs")
@click.option("--doc", type=click.Path(exists=True, path_type=Path), default=None, help="Open this document before listing paragraphs.")
@click.option("--limit", type=int, default=None, help="Max paragraphs to return.")
@click.pass_context
def list_paragraphs_command(ctx: click.Context, doc: Path | None, limit: int | None) -> None:
    """List document paragraphs."""
    session = _session_from_context(ctx)
    _open_if_requested(session, str(doc) if doc else None)
    _require_session_doc(session)
    rows = session.list_paragraphs(limit=limit)
    payload = {"ok": True, "action": "list-paragraphs", "paragraphs": rows}
    if ctx.obj.get("json_output", False):
        _emit(ctx, payload)
        return

    if not rows:
        click.echo("No paragraphs.")
        return
    for row in rows:
        click.echo(f"[{row['index']}] ({row['style']}) {row['text']}")


@cli.command("add-paragraph")
@click.argument("text")
@click.option("--doc", type=click.Path(exists=True, path_type=Path), default=None, help="Perform operation against this file and save it.")
@click.option("--style", default=None, help="Paragraph style name.")
@click.pass_context
def add_paragraph_command(ctx: click.Context, text: str, doc: Path | None, style: str | None) -> None:
    """Add a paragraph."""
    session = _session_from_context(ctx)
    used_doc = _open_if_requested(session, str(doc) if doc else None)
    _require_session_doc(session)
    session.add_paragraph(text=text, style=style)
    if used_doc:
        saved = session.save(doc)
    else:
        saved = session.path
    payload = {
        "ok": True,
        "action": "add-paragraph",
        "text": text,
        "style": style,
        "path": str(saved) if saved else None,
    }
    _emit(ctx, payload, text=f"Added paragraph. path={saved}")


@cli.command("add-heading")
@click.argument("text")
@click.option("--doc", type=click.Path(exists=True, path_type=Path), default=None, help="Perform operation against this file and save it.")
@click.option("--level", type=click.IntRange(0, 9), default=1, show_default=True)
@click.pass_context
def add_heading_command(ctx: click.Context, text: str, doc: Path | None, level: int) -> None:
    """Add a heading paragraph."""
    session = _session_from_context(ctx)
    used_doc = _open_if_requested(session, str(doc) if doc else None)
    _require_session_doc(session)
    session.add_heading(text=text, level=level)
    if used_doc:
        saved = session.save(doc)
    else:
        saved = session.path
    payload = {
        "ok": True,
        "action": "add-heading",
        "text": text,
        "level": level,
        "path": str(saved) if saved else None,
    }
    _emit(ctx, payload, text=f"Added heading. path={saved}")


@cli.command("add-table")
@click.option("--doc", type=click.Path(exists=True, path_type=Path), default=None, help="Perform operation against this file and save it.")
@click.option("--rows", type=click.IntRange(1, None), required=False, default=None, help="Row count for empty-table mode or minimum rows in structured mode.")
@click.option("--cols", type=click.IntRange(1, None), required=False, default=None, help="Column count for empty-table mode or column override in structured mode.")
@click.option("--header", "headers", multiple=True, help="Header cell value. Repeat for each header column.")
@click.option("--record", "record_values", multiple=True, help="Row values split by --delimiter (repeatable).")
@click.option("--delimiter", default="|", show_default=True, help="Delimiter used to parse each --record line.")
@click.option("--header-bold/--no-header-bold", default=False, help="Render header text in bold.")
@click.option(
    "--header-bg-color",
    default=None,
    help="Header background color from palette: main|secondary or exact palette hex.",
)
@click.option("--row-lines/--no-row-lines", default=False, help="Draw separator lines between rows.")
@click.option("--column-lines/--no-column-lines", default=False, help="Draw separator lines between columns.")
@click.option("--outer-border/--no-outer-border", default=False, help="Draw border around the table.")
@click.option(
    "--line-style",
    type=click.Choice(["single", "dashed", "dotted", "double", "thick"], case_sensitive=False),
    default="single",
    show_default=True,
    help="Border line style.",
)
@click.option("--line-size", type=click.IntRange(2, 96), default=8, show_default=True, help="Border width in eighths of a point.")
@click.option(
    "--line-color",
    default="main",
    show_default=True,
    help="Border color from palette: main|secondary or exact palette hex.",
)
@click.pass_context
def add_table_command(
    ctx: click.Context,
    doc: Path | None,
    rows: int | None,
    cols: int | None,
    headers: tuple[str, ...],
    record_values: tuple[str, ...],
    delimiter: str,
    header_bold: bool,
    header_bg_color: str | None,
    row_lines: bool,
    column_lines: bool,
    outer_border: bool,
    line_style: str,
    line_size: int,
    line_color: str,
) -> None:
    """Add an empty or data-populated table."""
    if len(delimiter) != 1:
        raise click.UsageError("--delimiter must be a single character.")

    records = _parse_records(record_values, delimiter)
    has_structured_data = bool(headers or records)
    if not has_structured_data and (rows is None or cols is None):
        raise click.UsageError("Provide --rows and --cols for empty-table mode, or use --header/--record.")
    if (header_bold or header_bg_color) and not headers:
        raise click.UsageError("--header-bold/--header-bg-color require at least one --header.")

    session = _session_from_context(ctx)
    used_doc = _open_if_requested(session, str(doc) if doc else None)
    _require_session_doc(session)
    table_shape = session.add_table(
        rows=rows,
        cols=cols,
        headers=list(headers),
        records=records,
        header_bold=header_bold,
        header_bg_color=header_bg_color,
        row_lines=row_lines,
        column_lines=column_lines,
        outer_border=outer_border,
        line_style=line_style.lower(),
        line_size=line_size,
        line_color=line_color,
    )
    if used_doc:
        saved = session.save(doc)
    else:
        saved = session.path
    payload = {
        "ok": True,
        "action": "add-table",
        "rows": table_shape["rows"],
        "cols": table_shape["cols"],
        "header_count": len(headers),
        "record_count": len(records),
        "header_bold": header_bold,
        "header_bg_color": header_bg_color,
        "row_lines": row_lines,
        "column_lines": column_lines,
        "outer_border": outer_border,
        "line_style": line_style.lower(),
        "line_size": line_size,
        "line_color": line_color,
        "path": str(saved) if saved else None,
    }
    _emit(
        ctx,
        payload,
        text=(
            f"Added table {table_shape['rows']}x{table_shape['cols']} "
            f"(headers={len(headers)}, records={len(records)}, "
            f"header_bold={header_bold}, header_bg_color={header_bg_color}, "
            f"row_lines={row_lines}, column_lines={column_lines}, outer_border={outer_border}). "
            f"path={saved}"
        ),
    )


@cli.command("frontpage-templates")
@click.pass_context
def frontpage_templates_command(ctx: click.Context) -> None:
    """List available frontpage templates."""
    session = _session_from_context(ctx)
    templates = session.list_frontpage_templates()
    payload = {"ok": True, "action": "frontpage-templates", "templates": templates}
    if ctx.obj.get("json_output", False):
        _emit(ctx, payload)
        return
    click.echo("\n".join(templates))


@cli.command("add-frontpage")
@click.option("--doc", type=click.Path(exists=True, path_type=Path), default=None, help="Perform operation against this file and save it.")
@click.option(
    "--template",
    "template_name",
    type=click.Choice(list(JsDocxSession.FRONTPAGE_TEMPLATES), case_sensitive=False),
    default="clean",
    show_default=True,
    help="Frontpage template preset.",
)
@click.option("--title", required=True, help="Frontpage title.")
@click.option("--subtitle", default=None, help="Optional subtitle.")
@click.option("--author", default=None, help="Optional author line.")
@click.option("--organization", default=None, help="Optional organization/institution line.")
@click.option("--date-text", default=None, help="Optional date text. Defaults to today's date.")
@click.option("--page-break/--no-page-break", default=True, help="Insert a page break after frontpage.")
@click.option("--set-core-title/--no-set-core-title", default=True, help="Update document core title metadata.")
@click.pass_context
def add_frontpage_command(
    ctx: click.Context,
    doc: Path | None,
    template_name: str,
    title: str,
    subtitle: str | None,
    author: str | None,
    organization: str | None,
    date_text: str | None,
    page_break: bool,
    set_core_title: bool,
) -> None:
    """Insert a template-driven frontpage at the beginning of the document."""
    session = _session_from_context(ctx)
    used_doc = _open_if_requested(session, str(doc) if doc else None)
    _require_session_doc(session)
    frontpage_meta = session.add_frontpage(
        template=template_name,
        title=title,
        subtitle=subtitle,
        author=author,
        organization=organization,
        date_text=date_text,
        include_page_break=page_break,
        set_core_title=set_core_title,
    )
    if used_doc:
        saved = session.save(doc)
    else:
        saved = session.path
    payload = {
        "ok": True,
        "action": "add-frontpage",
        "frontpage": frontpage_meta,
        "path": str(saved) if saved else None,
    }
    _emit(
        ctx,
        payload,
        text=(
            f"Inserted frontpage template '{frontpage_meta['template']}' "
            f"with title '{frontpage_meta['title']}'. path={saved}"
        ),
    )


@cli.command("add-bibliography-entry")
@click.argument("key")
@click.argument("reference")
@click.option("--url", default=None, help="Optional source URL.")
@click.option("--doc", type=click.Path(exists=True, path_type=Path), default=None, help="Perform operation against this file and save it.")
@click.pass_context
def add_bibliography_entry_command(
    ctx: click.Context,
    key: str,
    reference: str,
    url: str | None,
    doc: Path | None,
) -> None:
    """Add a bibliography entry and assign it a citation number."""
    session = _session_from_context(ctx)
    used_doc = _open_if_requested(session, str(doc) if doc else None)
    _require_session_doc(session)
    entry = session.add_bibliography_entry(key=key, reference=reference, url=url)
    if used_doc:
        saved = session.save(doc)
    else:
        saved = session.path
    payload = {
        "ok": True,
        "action": "add-bibliography-entry",
        "entry": entry,
        "path": str(saved) if saved else None,
    }
    _emit(
        ctx,
        payload,
        text=(
            f"Added bibliography entry [{entry['number']}] {entry['key']}. "
            f"path={saved}"
        ),
    )


@cli.command("list-bibliography")
@click.option("--doc", type=click.Path(exists=True, path_type=Path), default=None, help="Open this document before listing bibliography entries.")
@click.pass_context
def list_bibliography_command(ctx: click.Context, doc: Path | None) -> None:
    """List bibliography entries recognized by this harness."""
    session = _session_from_context(ctx)
    _open_if_requested(session, str(doc) if doc else None)
    _require_session_doc(session)
    entries = session.list_bibliography()
    payload = {"ok": True, "action": "list-bibliography", "entries": entries}
    if ctx.obj.get("json_output", False):
        _emit(ctx, payload)
        return

    if not entries:
        click.echo("No bibliography entries.")
        return
    for entry in entries:
        click.echo(f"[{entry['number']}] {entry['key']}: {entry['reference']}")


@cli.command("add-citation")
@click.argument("text")
@click.option("--source-key", "source_keys", multiple=True, help="Bibliography key to cite. Repeat for multi-source citations.")
@click.option("--style", default=None, help="Paragraph style name.")
@click.option("--doc", type=click.Path(exists=True, path_type=Path), default=None, help="Perform operation against this file and save it.")
@click.pass_context
def add_citation_command(
    ctx: click.Context,
    text: str,
    source_keys: tuple[str, ...],
    style: str | None,
    doc: Path | None,
) -> None:
    """Add a paragraph ending with in-text citation marker(s)."""
    normalized_source_keys = _require_source_keys(source_keys)

    session = _session_from_context(ctx)
    used_doc = _open_if_requested(session, str(doc) if doc else None)
    _require_session_doc(session)
    citation = session.add_citation(text=text, source_keys=normalized_source_keys, style=style)
    if used_doc:
        saved = session.save(doc)
    else:
        saved = session.path
    payload = {
        "ok": True,
        "action": "add-citation",
        "citation": citation,
        "path": str(saved) if saved else None,
    }
    _emit(
        ctx,
        payload,
        text=(
            f"Added citation {citation['citation_marker']} for keys "
            f"{', '.join(citation['source_keys'])}. path={saved}"
        ),
    )


@cli.command("cite-paragraph")
@click.option("--index", "paragraph_index", type=click.IntRange(0, None), required=True, help="Zero-based paragraph index to annotate.")
@click.option("--source-key", "source_keys", multiple=True, help="Bibliography key to cite. Repeat for multi-source citations.")
@click.option("--doc", type=click.Path(exists=True, path_type=Path), default=None, help="Perform operation against this file and save it.")
@click.pass_context
def cite_paragraph_command(
    ctx: click.Context,
    paragraph_index: int,
    source_keys: tuple[str, ...],
    doc: Path | None,
) -> None:
    """Append in-text citation marker(s) to an existing paragraph."""
    normalized_source_keys = _require_source_keys(source_keys)

    session = _session_from_context(ctx)
    used_doc = _open_if_requested(session, str(doc) if doc else None)
    _require_session_doc(session)
    citation = session.cite_paragraph(paragraph_index=paragraph_index, source_keys=normalized_source_keys)
    if used_doc:
        saved = session.save(doc)
    else:
        saved = session.path
    payload = {
        "ok": True,
        "action": "cite-paragraph",
        "citation": citation,
        "path": str(saved) if saved else None,
    }
    _emit(
        ctx,
        payload,
        text=(
            f"Updated paragraph {paragraph_index} with citation {citation['citation_marker']}. "
            f"path={saved}"
        ),
    )


@cli.command("set-core")
@click.argument("key")
@click.argument("value")
@click.option("--doc", type=click.Path(exists=True, path_type=Path), default=None, help="Perform operation against this file and save it.")
@click.pass_context
def set_core_command(ctx: click.Context, key: str, value: str, doc: Path | None) -> None:
    """Set a core document property."""
    session = _session_from_context(ctx)
    used_doc = _open_if_requested(session, str(doc) if doc else None)
    _require_session_doc(session)
    session.set_core_property(key=key, value=value)
    if used_doc:
        saved = session.save(doc)
    else:
        saved = session.path
    payload = {
        "ok": True,
        "action": "set-core",
        "key": key,
        "value": value,
        "path": str(saved) if saved else None,
    }
    _emit(ctx, payload, text=f"Set core property {key}. path={saved}")


@cli.command("undo")
@click.pass_context
def undo_command(ctx: click.Context) -> None:
    """Undo last mutation in current session."""
    session = _session_from_context(ctx)
    session.undo()
    summary = session.summary()
    payload = {"ok": True, "action": "undo", "summary": summary}
    _emit(ctx, payload, text=f"Undo complete\n{_format_summary(summary)}")


@cli.command("redo")
@click.pass_context
def redo_command(ctx: click.Context) -> None:
    """Redo last undone mutation in current session."""
    session = _session_from_context(ctx)
    session.redo()
    summary = session.summary()
    payload = {"ok": True, "action": "redo", "summary": summary}
    _emit(ctx, payload, text=f"Redo complete\n{_format_summary(summary)}")


if __name__ == "__main__":
    cli()
