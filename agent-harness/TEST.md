# Test Plan

This harness includes two test layers:

1. `test_core.py`
- validates `DocxSession` state transitions
- confirms mutating operations alter document state correctly
- verifies undo/redo stack behavior
- checks save/open round-trip and metadata writes

2. `test_full_e2e.py`
- exercises the installed CLI command via subprocess
- verifies one-shot JSON output and persisted `.docx` edits
- verifies default REPL path and in-session undo

## Validation Profiles

From `python-docx/agent-harness`:

### JS-default profile

```bash
python -m pip install -e .
npm install
DOCX_ENGINE=auto cli-anything-python-docx --json engine
python -m pytest -q cli_anything/python_docx/tests/test_core.py cli_anything/python_docx/tests/test_full_e2e.py
```

### Python-fallback profile

```bash
python -m pip install -e .[python]
npm install
DOCX_ENGINE=python cli-anything-python-docx --json engine
python -m pytest -q cli_anything/python_docx/tests/test_core.py cli_anything/python_docx/tests/test_full_e2e.py
```

Notes:

- In JS-default profile, tests that require `python-docx` parser/backends are skipped automatically when unavailable.
- In Python-fallback profile, marker-based citation behavior and Python-session compatibility remain covered.
