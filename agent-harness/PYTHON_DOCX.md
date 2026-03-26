# PYTHON_DOCX Harness

Target software: `python-docx`
Source path: `python-docx/`

## Scope

This harness exposes high-value DOCX operations as a stateful CLI with JS-first runtime behavior:

- create/open/save documents
- add paragraph, heading, and table content
- insert template-driven frontpages/title pages
- add bibliography entries and in-text citations for claim/data traceability
- generate Word footnote citations in JS mode and marker citations in Python mode
- update core metadata properties
- inspect summary and paragraph lists
- run with one-shot subcommands or interactive REPL
- emit machine-readable JSON output using `--json`
- support undo/redo in-session via in-memory `.docx` snapshots
- enforce a fixed visual palette (`#05206E` main, `#357AE9` secondary)

## Backend strategy

Primary runtime is the Node worker (`utils/js_engine/engine.mjs`) selected by default when
dependencies are available (`DOCX_ENGINE=auto`). The JS session supports one-shot `--doc` mode
and in-memory REPL/session mode. JS session `new` starts from a bundled default `.docx` template
(`templates/default.docx`) to avoid Python-side document construction requirements.

Python fallback mode remains available through `utils/python_docx_backend.py` and `DocxSession`.

Engine selection is controlled by `DOCX_ENGINE` (`auto`, `js`, `python`), where `auto`
prefers JS when Node dependencies are available. Per-invocation override is also available via
`--engine auto|js|python`.

## Session model

A `DocxSession` instance tracks:

- active in-memory document
- current path (if saved/opened)
- undo stack
- redo stack

Undo and redo operate on serialized `.docx` bytes captured before each mutating operation.
