# Agent instructions

Read this before changing anything in this repository.

## What this repository is

A converter that turns Microsoft Outlook Template files (`.oft`) into RFC-compliant `.eml` messages, preserving HTML bodies, plain-text alternatives, attachments, and inline images with their Content-ID references. It ships two entry points over one conversion core: a command-line script and a Tkinter GUI. Users run it on their own machines against their own files.

A change to `oft_to_eml_converter.py` reaches every consumer of both entry points. The GUI is a shell around it and holds no conversion logic.

## What this repository is not

It is not the macOS application. That is [`trsdn/oft-eml-converter-mac`](https://github.com/trsdn/oft-eml-converter-mac), a separate SwiftUI app with its own Python converter. The two are independent; a fix here does not reach there, and porting one to the other is a deliberate act.

It is also not an MSG parser. See [`docs/decisions/0001-delegate-msg-parsing.md`](docs/decisions/0001-delegate-msg-parsing.md).

## Layout

| Path | Purpose |
|---|---|
| `oft_to_eml_converter.py` | Conversion core and CLI entry point. All parsing and MIME assembly lives here |
| `oft_to_eml_gui.py` | Tkinter GUI. Presentation only, calls into the core |
| `run_gui.sh` | Launcher that creates `venv/` on first run, then starts the GUI |
| `tests/` | Pytest suite |
| `tests/sample_files/` | Fixtures for the suite |
| `requirements.txt` | Runtime and test dependencies |
| `docs/` | Architecture and decision records |

Generated, never hand-edit:

- `.github/badges/conformance.svg` — rendered from `.github/conformance.yml`
- `.github/stats/` — rendered by `.github/workflows/stats.yml`

Everything else is hand-maintained and editable.

## Setup

```sh
python3 -m venv venv
. venv/bin/activate
python -m pip install -r requirements.txt
```

Python 3.10 or newer. The supported range is pinned by the CI matrix in `.github/workflows/test.yml` and stated in the README; those two must agree.

`tkinter` is required for the GUI and cannot be installed with pip. On macOS use `brew install python-tk`, on Debian or Ubuntu `apt-get install python3-tk`.

## Run

```sh
python oft_to_eml_converter.py input.oft [output.eml]   # CLI
python oft_to_eml_gui.py                                # GUI
```

## Validate before proposing a change

This is the single command that must succeed:

```sh
python -m pytest tests/ -v
```

CI additionally runs `flake8`, `black`, `isort`, and `mypy`. Run these too when the change touches Python source, because the lint job is a required check:

```sh
flake8 . --count --select=E9,F63,F7,F82 --show-source --statistics
black --check .
isort --check-only .
```

A change that cannot be validated is not ready, however obviously correct it looks.

## Conventions

- The conversion core must stay importable without a display. The GUI imports the core, never the reverse — CI runs headless, and an import-time Tkinter dependency breaks it.
- CLI output is plain ASCII with no colour and no Unicode decoration. This is a deliberate accessibility property, not an oversight. Do not add colour, spinners, box-drawing characters, or emoji.
- The tool performs no network access at runtime. Adding a call that reaches the network contradicts what the README promises about data handling, and requires updating it in the same change.
- Follow PEP 8. Line length is 88, to match `black`.

## Do not do these

- Do not rewrite history, force push, or delete branches.
- Do not commit secrets, tokens, credentials, or personal data. `tests/sample_files/` must never contain a real message from a real mailbox.
- Do not create or move tags, publish releases, or change repository settings. The maintainer does that.
- Do not hand-edit generated files. Regenerate them instead; the paths and their generators are listed under [Layout](#layout).
- Do not implement MSG or OFT binary parsing by hand. The format is delegated to `extract-msg` on purpose.
- Do not add a runtime dependency without asking. The dependency surface is deliberately narrow.
- Do not add network access, telemetry, crash reporting, or analytics.

## Attribution

Agent-authored commits carry a `Co-authored-by:` trailer naming the agent. Changes reach `main` through a pull request; the required checks are `Test Suite` and `lint`. Automation does not merge its own work.
