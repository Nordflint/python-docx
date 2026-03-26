# CLI-Anything python-docx Harness

Stateful CLI harness for `python-docx` with one-shot commands, JSON output, and REPL-first workflow.

## Install

From `agent-harness/`:

```bash
python -m pip install -e .
npm install
```

Optional Python fallback engine support:

```bash
python -m pip install -e .[python]
```

## Usage

One-shot examples:

```bash
cli-anything-python-docx --json new ./demo.docx --title "Demo"
cli-anything-python-docx add-paragraph --doc ./demo.docx "Hello from CLI"
cli-anything-python-docx --json summary --doc ./demo.docx
cli-anything-python-docx frontpage-templates
cli-anything-python-docx add-frontpage --doc ./demo.docx --template clean --title "Operations Review" --subtitle "Q1 2026" --author "CLI Agent" --organization "Nordflint" --date-text "2026-03-26"
cli-anything-python-docx add-table --doc ./demo.docx --header Qty --header Id --header Desc --record "3|101|Spam" --record "7|422|Eggs" --record "4|631|Spam, spam, eggs, and spam"
cli-anything-python-docx add-table --doc ./demo.docx --header Qty --header Id --header Desc --record "3|101|Spam" --record "7|422|Eggs" --record "4|631|Spam, spam, eggs, and spam" --header-bold --row-lines --column-lines --outer-border --line-style single --line-size 8 --line-color main
cli-anything-python-docx add-bibliography-entry --doc ./demo.docx gov2025 "Government Statistics Annual Report (2025)" --url "https://example.com/report"
cli-anything-python-docx add-citation --doc ./demo.docx --source-key gov2025 "Employment increased by 4.2% in 2025."
cli-anything-python-docx cite-paragraph --doc ./demo.docx --index 3 --source-key gov2025
cli-anything-python-docx --json list-bibliography --doc ./demo.docx

# JS engine mode (footnote citations in .docx XML for --doc commands)
DOCX_ENGINE=js cli-anything-python-docx add-bibliography-entry --doc ./demo.docx gov2025 "Government Statistics Annual Report (2025)" --url "https://example.com/report"
DOCX_ENGINE=js cli-anything-python-docx add-citation --doc ./demo.docx --source-key gov2025 "Employment increased by 4.2% in 2025."
cli-anything-python-docx --engine js add-citation --doc ./demo.docx --source-key gov2025 "Employment increased by 4.2% in 2025."
```

REPL (default when no subcommand is provided):

```bash
cli-anything-python-docx
docx> new ./notes.docx --title "Notes"
docx> add-heading "Sprint" --level 2
docx> add-paragraph "Action item"
docx> undo
docx> save
docx> exit
```

## Commands

- `new [PATH] [--title TEXT]`
- `open PATH`
- `save [PATH]`
- `summary [--doc PATH]` (use global `--json` before the subcommand)
- `list-paragraphs [--doc PATH] [--limit N]` (use global `--json` before the subcommand)
- `add-paragraph [--doc PATH] [--style NAME] TEXT`
- `add-heading [--doc PATH] [--level N] TEXT`
- `frontpage-templates`
- `add-frontpage [--doc PATH] [--template NAME] --title TEXT [--subtitle TEXT] [--author TEXT] [--organization TEXT] [--date-text TEXT] [--page-break/--no-page-break] [--set-core-title/--no-set-core-title]`
- `add-table [--doc PATH] [--rows N] [--cols N] [--header TEXT ...] [--record TEXT ...] [--delimiter CHAR] [--header-bold] [--header-bg-color HEX] [--row-lines] [--column-lines] [--outer-border] [--line-style STYLE] [--line-size N] [--line-color COLOR]`
- `add-bibliography-entry [--doc PATH] [--url URL] KEY REFERENCE`
- `list-bibliography [--doc PATH]` (use global `--json` before the subcommand)
- `add-citation [--doc PATH] [--style NAME] --source-key KEY ... TEXT`
- `cite-paragraph [--doc PATH] --index N --source-key KEY ...`
- `set-core [--doc PATH] KEY VALUE`
- `undo`
- `redo`
- `engine` (use global `--json` before the subcommand)
- `repl`

`--doc` on mutating commands enables one-shot edits that auto-save back to that file.

Structured table mode (`--header`/`--record`) is useful for record-style inserts.
Each `--record` line is parsed by `--delimiter` (default `|`).
Border formatting can be enabled independently for row separators (`--row-lines`) and column separators (`--column-lines`), and optional outside borders (`--outer-border`).
Header formatting defaults to `main` background with white text, and can be customized with `--header-bold` and `--header-bg-color`.
Frontpage templates let the agent insert a predefined title-page layout with a single command.
The `corporate` frontpage template uses a blue (`main`) background with white text.
Bibliography entries are stored as numbered lines under a `Bibliography` heading.
In-text citation markers use those numbers, for example `[1]` or `[1, 2]`, so claims and data can be traced to their sources.

## JS Engine (Hybrid Migration Slice)

- Engine selection uses `DOCX_ENGINE` with these values:
  - `auto` (default): prefer JS when Node + dependencies are available, else fall back to Python.
  - `js`: force JS engine.
  - `python`: force Python engine.
- You can also override engine selection per invocation using global `--engine auto|js|python`.
- `python-docx` is now optional at install-time and only required for Python engine mode.
- Set `DOCX_ENGINE=js` to force selected commands through the Node engine.
- In `auto` or `js` mode, command execution runs through the JS-backed session (including REPL + one-shot flows).
- In JS mode, citations are written as Word footnote references (`word/footnotes.xml`) while preserving CLI JSON/text contracts.
- In JS mode, commands without `--doc` now run on a JS-backed in-memory session, including `new/open/save` and `undo/redo`.

## Palette

This harness enforces a fixed palette:

- `main`: `#05206E`
- `secondary`: `#357AE9`

Title/header fonts and table styling use these colors by default. Color options only accept `main`, `secondary`, or those exact hex values.
