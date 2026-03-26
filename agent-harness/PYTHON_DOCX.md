# PYTHON_DOCX Harness

Target software: `python-docx`
Source path: `python-docx/`

## Scope

This harness exposes high-value document operations from `python-docx` as a stateful CLI:

- create/open/save documents
- add paragraph, heading, and table content
- insert template-driven frontpages/title pages
- add bibliography entries and in-text citations for claim/data traceability
- optionally route citation/bibliography `--doc` commands through a JS engine (`DOCX_ENGINE=js`) for Word footnote output
- update core metadata properties
- inspect summary and paragraph lists
- run with one-shot subcommands or interactive REPL
- emit machine-readable JSON output using `--json`
- support undo/redo in-session via in-memory `.docx` snapshots
- enforce a fixed visual palette (`#05206E` main, `#357AE9` secondary)

## Backend strategy

The harness wraps the real `python-docx` APIs in `utils/python_docx_backend.py` and does not reimplement `.docx` behavior.

For hybrid migration, a Node worker (`utils/js_engine/engine.mjs`) can be selected for selected
commands in both one-shot `--doc` mode and JS-backed in-memory session mode. This allows footnote
citation output while preserving the existing CLI contract and REPL flow.

Engine selection is controlled by `DOCX_ENGINE` (`auto`, `js`, `python`), where `auto`
prefers JS when Node dependencies are available.

## Session model

A `DocxSession` instance tracks:

- active in-memory document
- current path (if saved/opened)
- undo stack
- redo stack

Undo and redo operate on serialized `.docx` bytes captured before each mutating operation.
