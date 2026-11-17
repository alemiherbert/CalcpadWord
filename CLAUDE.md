# Calcpad for Word: notes for Claude

**Who:** Alemi Herbert (Dokimos Consulting Engineers, dokimos.net, Uganda), a structural engineer. They write design calculations in desktop Word (64-bit); Tekla Tedds for Word and Tekla Structural Designer are also on their PC. They make the design decisions and prefer short, plain answers.

**This repository holds:**
- **the Calcpad for Word add-in:** `src/`, `tests/`, `tools/`, `installer/`;
- **the Dokimos calculation templates:** `templates/dokimos/`;
- **Calcpad's source:** `calcpad/`, a git submodule from Codeberg, pinned and never modified;
- **the first PC's working folder:** `scratchpad/`, for reference only.

**Read first:** `README.md` (usage, building, layout), then `PLAN.md` (decisions, findings, roadmap, the open issue). For the templates, read `templates/dokimos/Template source/README.md`.

This file also carries what Claude Code remembered on the first PC (its memory folder doesn't move with the project).

## Commands

```powershell
git submodule update --init                      # once, after cloning: Calcpad's source
dotnet test tests\CalcpadWord.Engine.Tests       # unit tests, no Word needed (168 on 2026-11-18)
tools\Build-Installer.ps1                        # out\installer\CalcpadWord-Setup-<version>.exe
out\installer\CalcpadWord-Setup-<version>.exe /quiet   # install it (per user) to test the real thing
tools\Register.ps1                               # or: register the development build directly
tools\Test-InWord.ps1                            # end-to-end test in a hidden Word (ask the user to keep Word closed)
tools\Build-Loader.ps1                           # only after changing src\CalcpadWord.Loader (downloads Zig once)
```

The version is set in `src\Directory.Build.props`. Bump it when the installer changes.

## Rules

- **Word automation:**
  - Always start a separate hidden Word (`New-Object -ComObject Word.Application`).
  - Identify it as the one new `WINWORD` process ID, and stop only that ID.
  - Never quit, kill or automate a Word that has a visible window or the user's documents. If a Word process wasn't started by your script, leave it alone, and don't sample its threads either.
  - Windows can put a Word window the user opens into a hidden test instance. `Test-InWord.ps1` detects that and leaves that Word running.
- **No MessageBox in a hidden Word:** it would block automation forever. The add-in already checks `Application.Visible`; keep it that way.
- **Per user, no administrator rights:** the installer, fonts and registration all go to HKCU and `%LOCALAPPDATA%`. The same applies to tools you install while developing; ask first.
- **Calcpad stays unmodified:** use only its public APIs (`ExpressionParser`, `MacroParser`, `OpenXmlWriter`, `Messages`). Updating it means moving the submodule to a newer commit and running all the tests.
- **Templates:**
  - Every new template is a new version, `<Name> vN.dotx`. Never overwrite an earlier version.
  - The user's hand-adjusted document is the base. Change its XML with a script and keep their positions exactly; don't regenerate it. The base is "Calculation title.docx", from which v2 and v3 were made.
  - New versions go to `templates/dokimos/`. Also copy them to the user's own Dokimos folder, if they have one outside the repository (on the first PC: `Desktop\Alemi Herbert\Dokimos`).
- **Dokimos template design (the user's decisions):**
  - A4, 0.6 in margins.
  - Libre Franklin, SemiBold for headings. They rejected Cambria, Bahnschrift and Tahoma.
  - The dokimos.net logo, with no company name, URL or tagline anywhere else.
  - A bottom-aligned cover, with the title, details and document control at the foot, and no header or footer.
  - A faint geometric corner pattern at the top right, inside the margins (office printers can't print edge to edge), with triangles on the mesh's 60° lines.
  - Contents numbered i, ii; the body numbered 1, 2 ("Sheet x of y" counts the body only). No footers.
  - Brand colours: ink `1C1710`, paper `F7F4EE`, blue `2E63C4`, green `2FA968`.
- **Calculations:** the user chose Eurocodes (EN 1992-1-1 for the RC beam), with recommended values unless stated. In block output, PASS uses `class="ok"` and FAIL an inline red (`style="color:#C00000"`). Calcpad's `class="err"` counts as a calculation error.
- **Windows PowerShell 5.1 pitfalls:**
  - A script with non-ASCII text (Greek letters) needs a UTF-8 BOM.
  - A COM collection returned through `InvokeMember` has no `.Item()`; walk collections through properties instead (`$doc.ContentControls.Item($i)`).
  - Inside a method call a comma separates arguments, so `$list.Add("{0}" -f $a, $b)` passes two arguments.
  - Quote arguments that contain commas, such as `'-Wl,--entry=DllMain'`.
  - `2>&1` on native programs turns stderr into errors.

## State on 2026-11-18 (version 0.8.0)

- **Done and installed on the first PC:**
  - inline and block evaluation, script conversion and background calculation;
  - a math font chooser, Cancel, and greying of out-of-date results;
  - **input fields** (form fields that set variables: reusable templates);
  - **one-file Setup:** the add-in, Calcpad's engine, a private .NET runtime started by the native loader DLL, and the fonts. It needs no .NET and no administrator rights.
- **Templates:**
  - `Dokimos Design Calculation v3.dotx` is the general template.
  - `Dokimos RC Beam Design (EC2) v1.dotx` is a complete beam calculation with input fields, PASS/FAIL checks and a summary. It needs add-in 0.8.0 or later.
  - Both are built and tested by the scripts in `templates/dokimos/Template source/`.
- **Open:**
  - An intermittent hang of the hidden test Word after the first background draw ("Open issue" in PLAN.md).
  - Setup isn't code-signed, so antivirus or policy can block it on other PCs. The error "Windows cannot access the specified device…" is fixed with Properties → Unblock, or by allowing it in Windows Security. Code signing (Azure Trusted Signing) was proposed but not set up.
  - Offered but not done: Setup could also install the Dokimos templates into Word's personal templates. A Tedds cross-check of the beam template is also pending: the Tedds Engineering Library wasn't installed or licensed on the first PC.
