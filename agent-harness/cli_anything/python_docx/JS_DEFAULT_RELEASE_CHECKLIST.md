# JS-Default Release Checklist

Target: ship the CLI with JS runtime as the default execution path while preserving Python fallback compatibility.

## Runtime and Packaging

- [x] `DOCX_ENGINE=auto` prefers JS when Node runtime and dependencies are present.
- [x] Global `--engine auto|js|python` override is available for one-shot and REPL flows.
- [x] `python-docx` dependency is optional (`pip install -e .[python]`) instead of required.
- [x] JS engine worker and default template are packaged with the harness.
- [ ] Validate clean install path on a fresh environment with only JS dependencies.
- [ ] Validate fallback install path with Python-only runtime (`DOCX_ENGINE=python`).

## Behavior and Parity

- [x] Core editing commands run through JS session in `auto/js` mode.
- [x] JS mode supports bibliography + footnote citations in `.docx`.
- [x] Paragraph listings in JS mode include rendered citation markers (`[n]` summary).
- [x] Python fallback citation behavior remains available for compatibility.
- [ ] Confirm command-by-command parity matrix for JS vs Python sessions.
- [ ] Decide whether any remaining Python-only behavior should be dropped before GA.

## Tests and CI

- [x] e2e tests cover JS auto preference and explicit engine override.
- [x] Python-specific tests are runtime-gated when Python backend is unavailable.
- [x] Test modules no longer hard-fail at import when `python-docx` parser is absent.
- [ ] Add CI matrix job: JS-default install (no `python-docx` extra).
- [ ] Add CI matrix job: Python fallback mode with `DOCX_ENGINE=python`.
- [ ] Add smoke test for `engine` command output in both matrix jobs.

## Docs and Release Ops

- [x] Harness README states JS-first defaults and Python fallback model.
- [x] Internal harness design doc updated for JS-first backend strategy.
- [ ] Add release notes section summarizing engine-mode differences.
- [ ] Communicate migration guidance for downstream agent prompts and scripts.
- [ ] Tag release candidate and run branch-to-main merge checklist.
