# CLI-Anything python-docx Harness

Stateful CLI harness for `python-docx` with one-shot commands, JSON output, and REPL-first workflow.

## Install

From `agent-harness/`:

```bash
python -m pip install -e .
```

## Usage

One-shot examples:

```bash
cli-anything-python-docx --json new ./demo.docx --title "Demo"
cli-anything-python-docx add-paragraph --doc ./demo.docx "Hello from CLI"
cli-anything-python-docx --json summary --doc ./demo.docx
cli-anything-python-docx add-table --doc ./demo.docx --header Qty --header Id --header Desc --record "3|101|Spam" --record "7|422|Eggs" --record "4|631|Spam, spam, eggs, and spam"
cli-anything-python-docx add-table --doc ./demo.docx --header Qty --header Id --header Desc --record "3|101|Spam" --record "7|422|Eggs" --record "4|631|Spam, spam, eggs, and spam" --header-bold --header-bg-color D9E1F2 --row-lines --column-lines --outer-border --line-style single --line-size 8 --line-color 000000
```

REPL (default when no subcommand is provided):

```bash
cli-anything-python-docx
python-docx> new ./notes.docx --title "Notes"
python-docx> add-heading "Sprint" --level 2
python-docx> add-paragraph "Action item"
python-docx> undo
python-docx> save
python-docx> exit
```

## Commands

- `new [PATH] [--title TEXT]`
- `open PATH`
- `save [PATH]`
- `summary [--doc PATH]` (use global `--json` before the subcommand)
- `list-paragraphs [--doc PATH] [--limit N]` (use global `--json` before the subcommand)
- `add-paragraph [--doc PATH] [--style NAME] TEXT`
- `add-heading [--doc PATH] [--level N] TEXT`
- `add-table [--doc PATH] [--rows N] [--cols N] [--header TEXT ...] [--record TEXT ...] [--delimiter CHAR] [--header-bold] [--header-bg-color HEX] [--row-lines] [--column-lines] [--outer-border] [--line-style STYLE] [--line-size N] [--line-color COLOR]`
- `set-core [--doc PATH] KEY VALUE`
- `undo`
- `redo`
- `repl`

`--doc` on mutating commands enables one-shot edits that auto-save back to that file.

Structured table mode (`--header`/`--record`) is useful for record-style inserts.
Each `--record` line is parsed by `--delimiter` (default `|`).
Border formatting can be enabled independently for row separators (`--row-lines`) and column separators (`--column-lines`), and optional outside borders (`--outer-border`).
Header formatting can be enabled with `--header-bold` and `--header-bg-color D9E1F2`.
