# Calcpad for Word: notes for Claude

**Who:** Alemi Herbert (Dokimos Consulting Engineers, dokimos.net, Uganda), a structural engineer. They write design calculations in desktop Word (64-bit); Tekla Tedds for Word and Tekla Structural Designer are also on their PC. They make the design decisions and prefer short, plain answers.

**This repository holds:**
- **the Calcpad for Word add-in:** `src/`, `tests/`, `tools/`, `installer/`;
- **the Dokimos calculation templates:** `templates/dokimos/`;
- **the Dokimos corporate document system:** `templates/dokimos/Document System/`, on the first PC only and kept off GitHub until the user decides (2026-11-20): never `git add` it. Its README holds the user's decisions and its state;
- **Calcpad's source:** `calcpad/`, a git submodule from Codeberg, pinned and never modified;
- **the first PC's working folder:** `scratchpad/`, for reference only.

**Where it lives:** on GitHub, [github.com/alemiherbert/CalcpadWord](https://github.com/alemiherbert/CalcpadWord) (`origin`, branch `main`), reached over SSH: `git@github.com:alemiherbert/CalcpadWord.git`. The local folder is named `CalcpadWord`, after the repository (it was `Plugin` until 2026-11-20). Pull before starting work on a PC, since both PCs push to it.

**Read first:** `README.md` (usage, building, layout), then `PLAN.md` (decisions, findings, roadmap, the open issue). For the templates, read `templates/dokimos/Template source/README.md`.

This file also carries what Claude Code remembered on the first PC (its memory folder doesn't move with the project).

## Commands

```powershell
git submodule update --init                      # once, after cloning: Calcpad's source
dotnet test tests\CalcpadWord.Engine.Tests       # unit tests, no Word needed (228 on 2026-11-20)
tools\Build-Installer.ps1                        # out\installer\CalcpadWord-Setup-<version>.exe
out\installer\CalcpadWord-Setup-<version>.exe /quiet   # install it (per user) to test the real thing
tools\Register.ps1                               # or: register the development build directly
tools\Test-InWord.ps1                            # end-to-end test in a hidden Word (ask the user to keep Word closed)
tools\Build-Loader.ps1                           # only after changing src\CalcpadWord.Loader (downloads Zig once)
templates\dokimos\Template source\beam\Check-WithTedds.ps1 -OutDir <dir> [-Recommended]   # Tedds' own beam calc, for cross-checks
gh release create v<version> out\installer\CalcpadWord-Setup-<version>.exe --title "Calcpad for Word <version>" --notes-file <notes>   # publish a build (after commit, tag and push)
```

The version is set in `src\Directory.Build.props`. **Public numbering starts at 0.1.0** (the user, 2026-11-20: "0.8x is premature"). The builds numbered 0.4 to 0.8 were internal, only on the user's PCs; history in PLAN.md keeps their numbers. **Versions move slowly** (the user, 2026-11-18): don't bump for small fixes. Batch changes into the next meaningful release, and make one new template version per piece of work, rebuilt in place until it's done.

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
  - **No guidance lines** (the user, 2026-11-20: "as you assume they are for submission"): no instructions to the designer in a template, and no "e.g." hints on the cover.
  - Brand colours: ink `1C1710`, paper `F7F4EE`, blue `2E63C4`, green `2FA968`.
- **Calculations:** the user chose Eurocodes (EN 1992-1-1 for the RC beam), with recommended values unless stated. Calcpad's `class="err"` counts as a calculation error, so a failing design must not use it.
  - A check ends in one line, Tedds style, the verdict and the reason in one colour and size: `'<p class="ok">PASS - reason</p>` (green), or `'<p style="color:#C00000;background-color:#FFEEEE">FAIL - reason</p>` (red).
  - Don't wrap the verdict in `<b>`: Calcpad leaves a formatted run uncoloured, and bold capitals look larger than the reason.
- **Snappy and succinct (the user, 2026-11-18):** input must feel instant, and templates say things once: no lines that echo an input, no repeated assumptions, short labels and verdicts. **Calculate automatically** (`DocumentSettings.Automatic`, stored as `auto=0` when off) **must be on by default, templates included** (the user, 2026-11-21; templates shipped with it off until then). Off works like Excel's manual mode: a changed input waits for Calculate; an expression written or edited is calculated and drawn on its own (Calcpad runs only as far as it), and what follows waits.
- **Block output (the user's decisions, made in `BlockOutput`):**
  - Tables use Word's Grid Table 1 Light, full width, left-aligned, with a tiny cell pad.
  - Plots fill the text width.
  - A picture alone on a line is centred.
  - The beam sketch has a light grey section.
- **Tekla Tedds** (installed on the second PC) has an API: see PLAN.md, "RC beam template v2". Never touch a Tedds the user has open. A script stops only the TeddsCalc process it started.
- **Windows PowerShell 5.1 pitfalls:**
  - A script with non-ASCII text (Greek letters) needs a UTF-8 BOM.
  - A COM collection returned through `InvokeMember` has no `.Item()`; walk collections through properties instead (`$doc.ContentControls.Item($i)`).
  - Inside a method call a comma separates arguments, so `$list.Add("{0}" -f $a, $b)` passes two arguments.
  - Quote arguments that contain commas, such as `'-Wl,--entry=DllMain'`.
  - `2>&1` on native programs turns stderr into errors.

## State on 2026-11-21 (version 0.1.1)

- **Released:** 0.1.0 is the first public release on GitHub (tag `v0.1.0`, with Setup attached), 2026-11-20. It is the internal 0.8.0 plus the second PC's block output and Recalculate-only mode.
- **0.1.1 (2026-11-21), committed, not yet released on GitHub:** the SVG drawing fix for Word 2019, out-of-date results greyed later and in grey, the buttons renamed (Calculate, Convert to Calcpad), Calculate automatically on by default with Excel-style manual mode, Lock and Unlock (two small buttons; templates ship with their calculation locked, inputs open), and the templates without guidance lines.
- **Done in the internal 0.8.0, which is installed on the first PC:**
  - inline and block evaluation, script conversion and background calculation;
  - a math font chooser, Cancel, and greying of out-of-date results;
  - **input fields** (form fields that set variables: reusable templates);
  - **one-file Setup:** the add-in, Calcpad's engine, a private .NET runtime started by the native loader DLL, and the fonts. It needs no .NET and no administrator rights.
- **Templates:**
  - `Dokimos Design Calculation v4.dotx` is the general template: v3 without guidance lines (2026-11-20). The calculation templates below are built on it; earlier versions are kept.
  - `Dokimos Psychrometric Chart v3.dotx`: moist air from two known properties (ASHRAE Fundamentals ch. 1) with a chart, built from `psychrometric\Psychrometric chart.cpd` and checked against ASHRAE's tables (`PsychrometricTests`). v1 was built on the first PC by an earlier script (2026-11-18) and is kept.
  - `Dokimos Duct Sizing v2.dotx`: replaces McQuay's DuctSizer (`D:\ductsizer\ductsizer\DuctSizer.exe`, runs on this PC). Round and rectangular ducts by friction rate, velocity or diameter (ASHRAE Fundamentals ch. 21), built from `duct\Duct sizing.cpd` and checked against DuctSizer (`DuctTests`). DuctSizer uses a roughness of about 0.21 mm; the template's default is ASHRAE's 0.09 mm.
  - `Dokimos Soil Stresses and Settlement v2.dotx`: Boussinesq stresses and settlement under a rectangular load varying across its width (after Calcpad's soil example), built from `soil\Soil stresses and settlement.cpd` and checked in `SoilTests`. About 7 s to calculate.
  - Templates don't say where their calculation came from (the user, 2026-11-20).
  - Calculation templates share `Template source\TemplateBuild.ps1` (hidden Word, sections, drop-downs, save).
  - `Dokimos RC Beam Design (EC2) v3.dotx` is the current beam calculation, modelled on Tedds and Calcpad's examples. It covers cover, fire, bending, limits, spacing, shear, crack control and deflection (by ratio and by calculation), with a sketch and a summary. It is cross-checked against Tedds. v1 and v2 are kept (v2 was built on the second PC). It needs add-in 0.1.0 or later.
  - All are built and tested by the scripts in `templates/dokimos/Template source/`.
- **First PC:** the original PC, with Office Professional Plus 2019 (Word 16.0.10417), not Microsoft 365. Add-in 0.1.1 is installed from Setup (2026-11-21). The Tedds Engineering Library isn't installed or licensed here. Git reaches GitHub over SSH with `~\.ssh\id_ed25519` (2026-11-20).
- **Second PC (2026-11-18):** the project arrived as a zip. Python and git are also in WSL; use Windows git for this repo. Word is Microsoft 365 x64. Add-in 0.8.0 is installed here.
- **Open:**
  - An intermittent hang of the hidden test Word after the first background draw ("Open issue" in PLAN.md).
  - Setup isn't code-signed, so antivirus or policy can block it on other PCs. The error "Windows cannot access the specified device…" is fixed with Properties → Unblock, or by allowing it in Windows Security. Code signing (Azure Trusted Signing) was proposed but not set up.
  - Offered but not done: Setup could also install the Dokimos templates into Word's personal templates.
  - Speed: a full recalculation of the v2 beam takes about 4 s (it was about 20 s; PLAN.md, "Speed"). The next speed-up would be drawing the PASS/FAIL blocks together, as equations already are (about 1.5 s of the draw).
