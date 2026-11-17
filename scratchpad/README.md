# Scratchpad (copied 2026-11-18)

This is the working folder from the development sessions on the first PC, kept for reference.

Nothing here is used by the build, the tests or the installer. The scripts still contain that PC's paths, so they may need editing before they run.

Everything that mattered was moved into the project:
- the template scripts are in `templates/dokimos/Template source`;
- the tests are in `tests/` and `tools/Test-InWord.ps1`;
- the findings are in `PLAN.md`.

## Contents

| Path | What it was |
|---|---|
| `template/` | Working copy of the Dokimos template work. It holds the generator (`build_template.py`), the Word finisher, the v2/v3 restructuring scripts, the unpacked original document (`unpacked-1/`), page renders and logs. The kept versions of these scripts are in `templates/dokimos/Template source`. |
| `logo-download/` | The Dokimos logo (`logo.svg`) and the page and stylesheet it came from (dokimos.net), used for the brand colours. |
| `probe/` | One-off probes of Word behaviour, made while building the add-in. Each folder has its script's log and pictures: blocks, conversion, context menus, Edit Source, Enter, fonts, greying, keys, and others. |
| `watch/` | Logs from watching the end-to-end test while chasing the intermittent hidden-Word hang (see "Open issue" in PLAN.md). |
| `Get-WaitChain.ps1`, `Get-ThreadSamples.ps1`, `stacks.txt` | Hang diagnostics: Windows wait-chain traversal of the test Word, and thread samples taken with dotnet-stack. |
| `Watch-E2E.ps1` | Runs the end-to-end test with a watchdog. |
| `Test-PrivateRuntime.ps1`, `corehost-trace.txt` | The proof that .NET's COM host uses an already-loaded `hostfxr.dll` and its runtime. This is the idea behind the loader (see the installer notes in PLAN.md). |
| `Test-Loader.ps1` | Calls `CalcpadWord.Loader.dll` directly, as COM does, against a private runtime. |
| `Check-Installed.ps1` | Checks an installed add-in in a hidden Word: connected, calculates, and where .NET was loaded from. |
| `Zoom-Page.ps1`, `zoom-page4.png` | Renders part of a page at high resolution (it showed that a Φ glyph problem was only in the low-resolution pictures). |
| `list_body.py` | Lists a .docx/.dotx body: paragraphs, styles and content-control tags. |
| `beam-build/`, `beam-test/`, `installed-check/` | Logs of the last beam template build and test runs, and of the installed add-in check. |

## Left out

These are copies of build output, and can be regenerated:
- `loader-test/` (97 MB): the add-in plus a private .NET runtime;
- `private-runtime/` (77 MB): a copy of the .NET runtime;
- `setup-check/` (11 MB): a Setup exe built without its payload;
- `fonts/` and `fira-download/`: copies of the fonts that are in `installer/fonts`.
