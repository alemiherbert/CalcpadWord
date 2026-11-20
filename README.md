# Calcpad for Word

Evaluates [Calcpad](https://codeberg.org/Proektsoft/Calcpad) in desktop Word for Windows (64-bit). Calculations become native Word equations inside ordinary Word text. Plots, maps and control flow are drawn by Calcpad in blocks. The whole document is one Calcpad worksheet, calculated top to bottom, so everything shares variables, functions and macros.

## Installing

Run **`CalcpadWord-Setup-<version>.exe`** and start Word: a **Calcpad** tab appears on the ribbon.
- **Everything is in that one file:**
  - the add-in;
  - Calcpad's calculation engine;
  - a private copy of the .NET runtime;
  - the math fonts and Libre Franklin.

  Calcpad and .NET don't need to be installed.
- **It installs for the current user only** and doesn't need administrator rights. The fonts go into the user's own font folder, as Windows' *Install* does.
- **To remove it,** use **Settings → Apps → Calcpad for Word → Uninstall**. The fonts stay installed.
- **For scripted installs,** use `/quiet` and `/uninstall`.

Requirements: **64-bit Windows** and **64-bit Word**. Setup refuses 32-bit Office, which can't load the add-in.

The setup is not code-signed, so on another computer Windows SmartScreen may warn: choose *More info → Run anyway*.

## Using it

**Pasting a Calcpad script:** paste it into Word, select it and press **Convert to Calcpad**. The script becomes Word content:

| In the script | In Word |
|---|---|
| `"Title`, `<h1>`…`<h3>` | Heading 1 |
| `<h4>`, `<h5>`, `<h6>` | Heading 2–4 |
| `'text 'expression' more text` | An ordinary paragraph, with the value inline |
| `<strong>`/`<b>`, `<em>`/`<i>`, `<u>`, `<sub>`, `<sup>` | That formatting |
| `<hr/>` | A rule line |
| `<!--'x = 1'-->` (hidden settings) | Hidden, takes no space |
| Lines with only an expression | A paragraph with the equation |
| `$Plot`, `$Map`, HTML tables and lists, `#if`, loops, `#def`, `#include` | A **Calcpad block**, drawn by Calcpad |
| `#round`, `#rad`, `#nosub`, `#val`, `#hide` … | Switches on the expressions that follow |
| `b = ? {300}` (Calcpad's input field) | An **input field**: the variable name, then a blue form field holding 300 (see below) |

A few keywords can't be reproduced this way: `#md`, `#format`, `#split`, `#phasor`, `#novar`, `#input` and `#pause`. To show PASS or FAIL in a block's output, give a line `class="ok"` (green) or an inline colour such as `style="color:#C00000"`. Calcpad's own `class="err"` counts as a calculation error. A script that uses them stays one block, and the status bar says why. **Block to Text** (ribbon, or right-click a block) converts an existing block the same way.

**Input fields make reusable templates.** An input field is an ordinary Word content control (plain text, combo box or drop-down list) that sets a Calcpad variable where it stands. Whoever uses the document types a number, or picks an item, and never sees Calcpad syntax.
- **Recalculation:** leaving a changed field recalculates the document.
- **Making fields:** write `name = ? {default}` in a Calcpad script and convert it (Convert to Calcpad). Calcpad's older saved values, after the script, also fill fields that have no default. A field's tag is `calcpad-input:name`, so you can also make one by hand (Developer tab → Plain Text Content Control → Properties → Tag).
- **Units:** with `calcpad-input:name:unit` a bare number gets the unit, e.g. `calcpad-input:L:m`. Without a unit only a number is accepted, as in Calcpad's own unitless templates.
- **Drop-down lists:** each item's *value* is the number the calculation uses, and its *display name* is what you see. For example, *C30/37* has the value 30, which sets f_ck. Values must be unique in a list, so write 0.3 and 0.30 if two items mean the same.
- **Wrong entries:** a value that isn't a number (a typo, or `0,85` with a comma) is shaded red. It is listed under **Errors** with a message such as "b: enter a number". Everything that doesn't depend on it still calculates.
- **Inside a Calcpad block** an input can't be a form field, so it keeps its default as an ordinary assignment.

**Writing your own expressions:**
- **Select and convert.** Type formulas, select them (or put the cursor in one), and press **Convert to Calcpad**. Each paragraph becomes one expression.
- **Inline `{= … }`.** Write `{= expression }` anywhere in a sentence. It is picked up on the next Calculate.

```
L = 6m
q = 12kN/m
M = q*L^2/8
The span is {=#val L} and the midspan moment is {=#val M|kNm}.
```

**Blocks** keep multi-line Calcpad together: `#if`/`#else`/`#end if`, `#for`/`#while`/`#repeat` … `#loop`, `#def` macros, `#include` (relative to the document's folder), `$Plot`/`$Map`, HTML tables. There are two ways to make one:
- the **Calcpad Block** button, which works on the selected paragraphs, or on an empty line to start typing;
- **fences**: a line with only `{=`, the script, then a line with only `}`.

Rules for blocks:
- A block must close every `#if`, loop and `#def` it opens.
- State such as `#hide` or `#round` set inside a block applies only to that block.
- Errors appear in red in the block's output, as in Calcpad.
- **How output looks in Word:**
  - Tables with `class="bordered"` take Word's *Grid Table 1 Light*, the full width of the text, left-aligned. A document that defines that style its own way keeps its version.
  - Plots fill the width of the text.
  - A picture alone on a line, such as an inline `<svg>` drawing, is centred. Give an SVG `alt="name.svg"` so that Word shows it as a vector picture.
- **Slow blocks are drawn once.** A block whose inputs (everything above it, and its own lines) haven't changed is not drawn again. Editing text or expressions *below* a long-running block, such as one with integral maps, is instant. Changing something above it recalculates it.

The Calcpad tab has five groups, by job: **Calculation**, **Write**, **Errors**, **Protect** and **Document settings**.

| Ribbon (Calcpad tab) | Shortcut | What it does |
|---|---|---|
| **Calculation:** Calculate | **Ctrl+Alt+F9** | Calculates everything: main text, text boxes, footnotes and endnotes, then headers and footers |
| Calculate automatically | | On by default: leaving a changed input or expression recalculates (see below). Saved in the document |
| Cancel | **Esc**, while a calculation runs | Stops a background calculation or a math font download. Results keep their old values until you calculate again |
| **Write:** Convert to Calcpad | **Alt+Shift+Enter** | Turns the selection (or the current paragraph) into Calcpad (see above), then calculates. On a result, the shortcut shows its source instead |
| Calcpad Block | | Keeps the selected paragraphs (or an empty line) together as one block |
| Block to Text | right-click a block | Turns a block into Word text, as for a pasted script |
| Edit Source | **double-click** a result, or right-click → **Edit Calcpad Source** | Shows the source so you can change it. **Click outside to calculate** |
| **Errors:** Errors (n) | | Lists what failed. Pick one to go to it |
| Next Error | **Ctrl+Alt+F8** | Moves to the next error and shows its message in the status bar |
| **Protect:** Lock, Unlock | | Lock the Calcpad expressions and blocks in the selection (Ctrl+A for all): they still calculate, but can't be opened, converted or deleted. Input fields stay open. A guard against accidents, not a password: Unlock lifts it. The Dokimos calculation templates ship locked |
| **Document settings:** Decimals, Angles, Substitute values | | How results are written. Saved in the document |
| Math font, Plain Word math | | How equations look (see below). Saved in the document |

Ribbon KeyTips also work: Alt, C, P, then the letter shown on the button.

**Switches** go in front of an inline expression and affect only that expression:

| Switch | Effect |
|---|---|
| `#equ` | Equation, substituted values and result (the default) |
| `#val` | Result only, with units. Use it for numbers in sentences. Unlike Calcpad's `#val`, units are kept |
| `#noc` | Show the expression without calculating it |
| `#hide` | Calculate silently. The source stays in the document as hidden text |
| `#nosub` / `#varsub` | Turn substitution of values off or on |
| `#round n` | Use n decimals (0–15) |
| `#deg` / `#rad` / `#gra` | Angle units |

**Math font.** The **Math font** list on the Calcpad tab sets the font for every part of the equations in this document: variables, numbers, operators and units. It is stored in the document as Word's own math font. The list is in alphabetical order. It always offers these:

| Font | Look | Availability |
|---|---|---|
| Asana Math | Palatino-like | Free: SIL OFL |
| Cambria Math | Word's default | Comes with Windows and Office |
| Euler Math (AMS Euler) | Upright, handwritten style | Free: SIL OFL |
| Fira Math | Sans serif, matching Fira Sans | Free: SIL OFL |
| Latin Modern Math | LaTeX (Computer Modern) | Free: GUST font licence |
| Libertinus Math | Linux Libertine | Free: SIL OFL |
| Noto Sans Math | Sans serif | Free: SIL OFL |
| STIX Two Math | Times-like, scientific publishing | Free: SIL Open Font License |
| TeX Gyre Pagella Math | Palatino-like | Free: GUST font licence |
| XITS Math | Times-like (STIX-based) | Free: SIL OFL |

It also lists:
- **every other math font installed on the computer**, found automatically: any installed font with an OpenType MATH table, such as a commercial one like Minion Math;
- **the document's own math font**, if it is none of these.

Setup installs all the free fonts. If one has been removed since, the list marks it *(not installed)*. Choosing it asks first, then reinstalls it for you only, as Windows' own *Install* does: no administrator rights, and no restart needed. It reinstalls from the copy bundled with the add-in; a font that isn't bundled is downloaded from CTAN or Google Fonts. The file is checked to be that font with an OpenType MATH table before it is installed.

**Long calculations run in the background.** If Calcpad takes more than a quarter of a second, Word stays usable. The status bar says *calculating…* and the results appear when they are ready.
- If you change Calcpad text meanwhile, it calculates again.
- A newer calculation replaces the running one.
- **Cancel** (or **Esc**) stops it. Esc is only taken from Word while a calculation runs.
- **Results that may change are greyed meanwhile:** a light grey background, washed-out plots, and *Out of date:* in front of their title. Only results from the first new or edited expression on are greyed, since the worksheet runs top to bottom. When the results come in, the grey goes. A result that turned out unchanged is not redrawn: only its grey is removed. After **Cancel** the grey stays, also in the saved document, until the next Calculate.
- Results that arrive late are one Undo step of their own.
- Calcpad is loaded when Word starts, so the first calculation isn't slowed down.

**Calculate automatically** (Calculation group, saved in the document) is on by default, also in the Dokimos templates: leaving a changed input field or expression recalculates. Turn it off for fast input in a long calculation. It then works like Excel's manual mode:
- leaving a changed field changes nothing on the page, and the Calculate button reads *Calculate (out of date)* until you click it (Ctrl+Alt+F9);
- an expression you write or edit is calculated on its own when you leave it (Calcpad runs only as far as it), so it never stays open as source; what follows it waits for Calculate.

The setting is saved with the document.

How the add-in behaves:
- **Results are read-only.** Double-click one (or use Edit Source) to change it. When you click outside, it is calculated again, along with everything after it.
- **Keys while editing an expression's source:**
  - **Shift+Enter** adds another expression line.
  - **Enter** works like normal Word: it finishes, calculates, and moves on to a new paragraph.
  - Inside a block's script, both keys add lines.
- **Errors** turn into red, editable source. The message shows as the control's title and in the status bar. Fix the text and click outside.
- **Copy and paste.** A copied result keeps its source, even into another document.
- **Undo.** One Undo reverts a whole Convert to Calcpad or Calculate.
- **Small and large numbers** show as 5.7 × 10⁻⁹: the × sits on the line and the exponent has a real minus sign. Calcpad itself writes a raised × and a hyphen.
- **Word's typing substitutions are undone automatically:** en dashes become minus signs, ½ becomes (1/2), and smart quotes and non-breaking spaces are straightened.

**Turn this Word option off:** File → Options → Proofing → AutoCorrect Options → AutoFormat As You Type → **"\*Bold\* and \_italic\_ with real formatting"**. Otherwise Word can delete `*` and `_` from formulas as you type.

**Limits:**
- Inline expressions can't contain quotes or plots; those go in blocks.
- Calcpad stops a worksheet after 40 errors. Anything after that is marked "Not calculated".
- Comments in Word's review pane are not scanned.

## Building and testing

Requirements: Windows, the .NET 10 SDK, and 64-bit desktop Word (tested with Office 2019). Visual Studio is not needed.

**Setting up a development PC:**
1. **Install the tools:** Git for Windows, the .NET 10 SDK, and 64-bit desktop Word. The build uses Office's Word interop assemblies from `C:\Windows\assembly\GAC_MSIL`; Office installs them unless its ".NET Programmability Support" was left out. Python 3 is needed only for the Dokimos template scripts.
2. **Connect to GitHub over SSH** (once per PC): make a key with `ssh-keygen -t ed25519`, add the contents of `~\.ssh\id_ed25519.pub` at [github.com/settings/ssh/new](https://github.com/settings/ssh/new), then check it with `ssh -T git@github.com`. Git, and VS Code through it, then uses the key.
3. **Get the repository** from [github.com/alemiherbert/CalcpadWord](https://github.com/alemiherbert/CalcpadWord), into a folder named `CalcpadWord`: `git clone --recurse-submodules git@github.com:alemiherbert/CalcpadWord.git`. `--recurse-submodules` also fetches Calcpad's source, which is pinned to the tested commit and downloaded from Codeberg; in a copy cloned without it, run `git submodule update --init`.
4. **Set who you commit as** (once per PC): `git config --global user.name "Alemi Herbert"` and `git config --global user.email "..."`.
5. **Check it works:** run `dotnet test tests\CalcpadWord.Engine.Tests`, then `tools\Build-Installer.ps1`. Install the result with `out\installer\CalcpadWord-Setup-<version>.exe`, then run `tools\Test-InWord.ps1` with Word closed.
6. **Using Claude Code:** open the folder. `CLAUDE.md` gives it the project's context and rules.

```powershell
dotnet test tests\CalcpadWord.Engine.Tests      # unit tests (no Word needed)
tools\Build-Installer.ps1                       # out\installer\CalcpadWord-Setup-<version>.exe
tools\Build-Loader.ps1                          # rebuild the native loader (only after changing it)
tools\Register.ps1                              # development: build and register the add-in directly (current user)
tools\Test-InWord.ps1                           # end-to-end test in a separate hidden Word; never touches your open Word
tools\Unregister.ps1 [-RemoveFiles]             # development: remove what Register.ps1 did
```

- **The version** is set once, in `src\Directory.Build.props`.
- **What Setup contains:** `Build-Installer.ps1` bundles the following.
  - **The add-in.** Of Calcpad, only its engine goes in: `Calcpad.Core` and `Calcpad.OpenXml`, without the Bulgarian and Chinese messages. The desktop app, CLI, help and examples are left out.
  - **The .NET runtime.** The newest .NET 10 runtime installed on the build computer is copied into `runtime\`. The add-in then always runs on that copy, even where .NET is installed.
  - **The fonts.** These come from `installer\fonts`. `fonts.txt` lists each file with the name Windows registers it under, and `FONT-NOTICES.txt` has their copyrights and licences.
  - **Setup itself.** It is self-contained and trimmed, so it runs without .NET.
- **The loader:** `src\CalcpadWord.Loader` contains the native DLL that Word loads. .NET's COM host only looks for a machine-wide .NET, but it uses a `hostfxr` that is already loaded. So the loader loads `runtime\host\fxr\<version>\hostfxr.dll`, then hands over to the add-in's COM host.
  - The built DLL is kept next to its source, so building the installer needs no C compiler.
  - `Build-Loader.ps1` rebuilds it with Zig's C compiler. It downloads Zig once to `%LOCALAPPDATA%\CalcpadWord-build` and checks it against the published SHA-256.
  - `Register.ps1` registers the COM host directly instead, using the installed .NET.
- **Register.ps1 and the installer** register the same add-in, and whichever ran last is the one Word loads. Each build goes into its own folder, so neither needs Word closed: a running Word keeps the version it loaded until you restart it.
- **Log:** `%LOCALAPPDATA%\CalcpadWord\addin.log`. The setup's log is `%TEMP%\CalcpadWord-Setup.log`.
- **Test output:** `Test-InWord.ps1` writes test documents and pictures of the rendered pages to `out\e2e`. Its math-font step installs the free math fonts for the current user.
- **Keep Word closed while the end-to-end test runs.** Windows may put a Word window you open into the test's hidden Word. The test then fails, says so, and leaves that Word running for you instead of closing it.

## Layout

| Path | What it is |
|---|---|
| `src/CalcpadWord.Engine` | The document-as-worksheet evaluator, the script-to-Word converter, `{= }` and fence scanning, switches, Word markup. No Word dependency |
| `src/CalcpadWord.AddIn` | The COM add-in: ribbon, shortcuts, Word events, content controls, rendering |
| `src/CalcpadWord.Setup` | The per-user installer |
| `src/CalcpadWord.Loader` | The native DLL Word loads: it starts the bundled .NET runtime, then the add-in (C, built by `tools/Build-Loader.ps1`) |
| `src/Shared` | Shared by the add-in and the installer: the add-in's identity (COM class, ProgID) and per-user font installation |
| `installer/fonts/` | The fonts Setup installs, with `fonts.txt` (registry names) and `FONT-NOTICES.txt` (copyrights and licences) |
| `tests/CalcpadWord.Engine.Tests` | Unit tests, including Calcpad's soil stress example (`Fixtures/`) |
| `tools/` | Installer build, development install and uninstall, end-to-end test |
| `installer/THIRD-PARTY-NOTICES.txt` | Licences of the bundled open-source components (shipped with the add-in) |
| `calcpad/` | Calcpad source, unmodified: a git submodule ([codeberg.org/Proektsoft/Calcpad](https://codeberg.org/Proektsoft/Calcpad), MIT) |
| `templates/dokimos/` | The Dokimos design calculation templates (every version), and `Template source/` with the scripts that build and test them |
| `scratchpad/` | The first PC's working folder: probes, diagnostics and logs, kept for reference (see its README) |
| `PLAN.md` | Design notes, findings and roadmap |
| `CLAUDE.md` | Context and working rules for Claude Code |
