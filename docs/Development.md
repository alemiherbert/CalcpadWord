# Development

Needs Windows, the .NET 10 SDK and 64-bit Word (tested with Office 2019). No Visual Studio.

## Set up a PC

1. Install Git, the .NET 10 SDK and Word, with Office's ".NET Programmability Support". Python 3 only for the template scripts.
2. SSH key for GitHub: `ssh-keygen -t ed25519`, add `~\.ssh\id_ed25519.pub` at [github.com/settings/ssh/new](https://github.com/settings/ssh/new), check with `ssh -T git@github.com`.
3. `git clone --recurse-submodules git@github.com:alemiherbert/CalcpadWord.git` (into `CalcpadWord`). Cloned without the flag: `git submodule update --init`.
4. `git config --global user.name` and `user.email`.
5. Run the tests and build below. Claude Code reads `CLAUDE.md`.

## Commands

```powershell
dotnet test tests\CalcpadWord.Engine.Tests   # unit tests, no Word
tools\Build-Installer.ps1                    # out\installer\CalcpadWord-Setup-<version>.exe
tools\Register.ps1                           # register the development build
tools\Unregister.ps1 [-RemoveFiles]          # undo Register.ps1
tools\Test-InWord.ps1                        # end-to-end test in a hidden Word (keep Word closed)
tools\Build-Loader.ps1                       # only after changing the loader
```

- Version: `src\Directory.Build.props`.
- Register.ps1 and Setup register the same add-in; the last one wins. Each build gets its own folder, so Word can stay open and keeps the version it loaded.
- Test-InWord writes documents and page pictures to `out\e2e`. A Word you open during it can join the test's Word; the test then leaves it running.

## What Setup bundles

- The add-in, and Calcpad's engine only (`Calcpad.Core`, `Calcpad.OpenXml`).
- The newest .NET 10 runtime on the build PC, in `runtime\`.
- The fonts in `installer\fonts` (`fonts.txt`: registry names; `FONT-NOTICES.txt`: licences).
- The loader (`src\CalcpadWord.Loader`), the native DLL Word loads: it loads the bundled `hostfxr`, then the add-in's COM host. The built DLL is in the repository; `Build-Loader.ps1` rebuilds it with Zig (downloaded once, SHA-256 checked).

## Layout

| Path | What |
|---|---|
| `src/CalcpadWord.Engine` | Worksheet evaluator, script converter, scanning, Word markup. No Word dependency |
| `src/CalcpadWord.AddIn` | The COM add-in: ribbon, events, rendering |
| `src/CalcpadWord.Setup` | Per-user installer |
| `src/CalcpadWord.Loader` | Native loader (C) |
| `src/Shared` | Add-in identity, font installation |
| `installer/` | Fonts and third-party notices |
| `tests/CalcpadWord.Engine.Tests` | Unit tests |
| `tools/` | Build, register, end-to-end test |
| `calcpad/` | Calcpad source, unmodified submodule ([Codeberg](https://codeberg.org/Proektsoft/Calcpad), MIT) |
| `templates/dokimos/` | Dokimos calculation templates and their build scripts |
| `scratchpad/` | The first PC's working folder, for reference |
| `PLAN.md` | Decisions, findings, roadmap |
| `CLAUDE.md` | Context for Claude Code |

[Home](../README.md)
