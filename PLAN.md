# Calcpad for Word: design notes and roadmap

Usage and build instructions are in [README.md](README.md). This file records the decisions, the findings behind them, and what's next.

## Decisions (2026-11-17)

- **Platform:** desktop Word on Windows only. That rules out Office.js.
- **Marking Calcpad text:** both ways. You can select text and press Evaluate, or write `{= … }` inline in a sentence.
- **Result style:** both, as a per-document setting. One is Calcpad's look (blue variable names); the other is plain Word math.

## Architecture (built and verified)

**A .NET 10 in-process COM add-in.**
- Word loads `CalcpadWord.AddIn.comhost.dll` (registered under HKCU). The comhost starts .NET 10 and creates `CalcpadWord.Connect`, which implements `IDTExtensibility2` and `IRibbonExtensibility`.
- The add-in references `Calcpad.Core` directly.
- It builds with the `dotnet` CLI. The Word and Office interop types are embedded from the copies Office installs in the GAC.
- Verified on Word 2019, build 16.0.10417: it loads alongside the VSTO add-ins (Tedds, PDFelement), and Word exits cleanly afterwards.
- `Register.ps1` installs each build into its own folder (`AddIn\v<date-time>`) and points the registration at it. A running Word keeps the files it loaded, so installing never needs Word closed. Old folders are deleted when they are no longer locked.

**Document model.**
- Each inline expression is an inline rich-text **content control**. Each block is a block-level one that wraps whole paragraphs.
- Its **Tag** holds everything: `calcpad:{r|e|x}:{fingerprint|-}:{source}` for inline expressions and `calcpad-block:…` for blocks.
  - Word keeps tags of at least 140,000 characters, with line breaks and any Unicode, so each control is self-contained. Copy and paste (even between documents) and undo just work.
- Document settings live in the document variable `Calcpad.Settings`.

**Recalculation: the document is one Calcpad worksheet** (`Worksheet.cs`).
1. Convert every `{= … }`, and every fenced block (`{=` alone on a line … `}` alone on a line), into a control.
2. Build one script. Each inline expression is one line, wrapped in keyword lines for its switches (`#round 3` … `#round default`). Each block contributes its lines, followed by a reset to the document defaults: `#show #equ #varsub|#nosub #wrap #round default #format default #md off #complex #deg|#rad|#gra`.
3. Put a marker comment line (`'CPWITEM<n>CPWITEM`) before every item.
4. Expand macros with `MacroParser`, then run `ExpressionParser` once.
5. Split the HTML at the marker paragraphs.
   - Splitting at markers, rather than by `line-N` IDs, stays exact when macros change the number of lines.
   - An inline item takes its equation's OMML (`eq-K` → `OpenXmlExpressions[K]`).
   - A block's HTML (equation IDs renumbered from 0, so the fingerprint is stable) is converted by Calcpad's own `OpenXmlWriter` to .docx. That becomes Flat OPC (with `sectPr` removed) and goes in through `InsertXML`.
6. Blocks are checked before running, and kept out of the script if they fail the check:
   - every `#if`, loop and `#def` they open must be closed;
   - no `#pause` or `#input`;
   - the last line must not end in a continuation character (`;|&@:({[`), which Calcpad would join with the next item.
7. Calcpad's messages are fixed to English (`Messages.Culture`), so error text can be parsed. The "on line [n]" part is stripped, because script line numbers mean nothing in Word.
8. Redraw only the controls whose fingerprint changed. Everything happens inside one custom undo record, with screen updating off.

**Rendering.** All rendering goes through `Range.InsertXML` (Flat OPC), including error text and source text. Setting `Range.Text` on a control that holds an equation would put the text *inside* the math zone.

**Control states** (the letter in the tag): `r` result, `e` source being edited, `x` failed source.
- Editing and failed controls are read back from their text when evaluated. A block's source is shown and read back as one paragraph per line.
- A block that ran but has Calcpad errors stays `r` and shows its output with the errors in red. Its title starts with "Error:", so it is listed under Errors.
- **Keys in an inline source** (the user's choice, 2026-11-17): **Shift+Enter** adds an expression line, and **Enter** works like normal Word (finish and move on).
  - Shift+Enter is Word's own line break (`\v`), which is allowed even inside a sentence.
  - When a source with line breaks or paragraph marks is evaluated (not while the cursor is still in it), it is unwrapped with `ContentControl.Delete(false)`. A range around it, boundaries included, follows the text, and each non-empty line is wrapped again in place as its own expression.
  - Enter is taken from Word by the keyboard hook only while `_enterLeavesControl` is true.
    - That flag is refreshed on `WindowSelectionChange`, on document switches, and after every command. The hook must not call Word itself.
    - The deferred action (`LeaveWithEnter`) re-checks where the cursor is, and falls back to `Selection.TypeParagraph()`, so a stale flag can't eat an ordinary Enter.
  - What Enter does on an inline expression: if nothing follows it on the line, insert a paragraph after it and move there; otherwise move just past it. Leaving the control then triggers the usual exit evaluation.
  - On a block's output, Enter starts a new paragraph after the block. In a block's script, Enter adds a line.
- Leaving an edited control (`ContentControlOnExit`, subscribed per document) re-evaluates it. A failed control is re-evaluated only if its text changed.

**Events and shortcuts** (`WordEvents.cs`, `KeyboardShortcuts.cs`, `UiThread.cs`):
- Work triggered by an event or a key is deferred with a one-shot Win32 timer, so it runs on Word's UI thread after Word finishes the click or keystroke.
- A recalculation triggered by leaving a control does not redraw the source the cursor moved into.
- Double-click is handled by `WindowBeforeDoubleClick`, which cancels Word's own handling.
- Right-click menus are ribbon `<contextMenus>` generated in `Connect.ContextMenusXml()`:
  - "Edit Calcpad Source" and "Evaluate Calcpad" are added to `ContextMenuEquation`, which is what a result shows, and to the text, table, list, heading, header, footer and spelling menus.
  - A `getVisible` callback shows the item that fits the control under the cursor.
  - `WindowBeforeRightClick` invalidates those items, so they are re-checked on every right-click.
  - The menu IDs were confirmed with `CommandBars.GetLabelMso`, which throws on unknown IDs. The end-to-end test re-checks them.
- Calcpad source text is marked `noProof`, so Word's spelling underline doesn't compete with the error underline.
- Shortcuts come from a keyboard hook on Word's UI thread only. The chords were chosen with `Application.FindKey` as unassigned in Word: Alt+Shift+Enter, Ctrl+Alt+F9 and Ctrl+Alt+F8.
  - Ctrl+Alt+letter was avoided because it means AltGr on European keyboards.
  - Alt+letter was avoided because it opens the ribbon KeyTips.
  - Keys are never swallowed.

**Stories.** Evaluation order is: main text, text boxes, footnotes, endnotes, then headers and footers. Linked headers are visited once (by content-control ID). All ranges are built from their own story (`Duplicate` + `SetRange`), not from `Document.Range`.

## Calcpad engine findings

- **The add-in uses only public Calcpad APIs:** `MacroParser`, `ExpressionParser` (with `HtmlResult` and `OpenXmlExpressions`), `OpenXmlWriter` and `Messages`. The clone stays unmodified, and the earlier reflection on `MathParser` internals is gone.
- `#round default`, `#format default` and `#md off` reset Calcpad's state.
  - `#varsub` and `#nosub` switch substitution. The document's "Substitute values" setting is applied this way, so a per-expression `#varsub` can override it.
  - `#complex` turns phasor display off.
- **Calcpad stops a worksheet after 40 errors.** Items after that point get "Not calculated".
- `MacroParser` reports its errors as text lines in its output. These are removed and attributed to the item whose marker they follow.
- `OpenXmlWriter` doesn't need Calcpad's CSS template. It handles `eq`, `err`, `img`, `svg`, tables and `h3` itself.
  - Its heading styles are based on `Normal`, so they pick up the user's document styles. Only the "Calcpad" and "Calcpad Bordered" table styles are added.
  - Plots arrive as embedded pictures.
- Rounding: Calcpad keeps significant figures for small values. With 2 decimals, 0.714 shows as `0.714`.
- `ToXml()`/`OpenXmlExpressions` produce OMML (Word's equation markup). Calcpad hard-codes 10–11 pt for units; dropping `w:sz` makes equations follow the paragraph.

## Environment notes

- Office Professional Plus 2019 (perpetual), x64. .NET 10 SDK. No Visual Studio, no Node.
- Automating Word: `New-Object -ComObject Word.Application` starts a **separate** hidden instance, which also loads COM add-ins.
  - Always identify it by PID and never `Quit` a Word you didn't start.
  - Call methods with by-ref arguments (`Quit`, `Close`, `SaveAs2`, `Open`, `Undo`) late-bound, and unwrap `PSObject` arguments.
- `ExportAsFixedFormat` (PDF) hangs in the hidden instance, probably because of the PDFelement add-in. For pictures, use `Range.EnhMetaFileBits` instead.
- `CommandBars.GetImageMso` can't draw any icon in a hidden Word (E_UNEXPECTED, even for "Bold"), but an unknown icon name fails differently ("Value does not fall within the expected range"). The end-to-end test uses that difference to check the ribbon icons.
- **Calls from other processes arrive on COM worker threads (MTA), not Word's UI thread (STA).** .NET objects are free-threaded, so `COMAddIn.Object` calls land on RPC threads. The add-in moves them onto the UI thread (`UiThread.Invoke`), except `IsWorking`/`LastStatus`, which only read fields.
- **Diagnosing hangs:**
  - `dotnet-stack report -p <Word PID>` (a global dotnet tool, installed 2026-11-17 with the user's approval) shows the add-in's managed stacks. A UI thread missing from the report is idle in Word's native code.
  - Creating `%LOCALAPPDATA%\CalcpadWord\trace.on` before Word starts turns on TRACE lines (UI work, events) in the log.
- **A safety timer (250 ms) on the add-in's message window** runs queued work if its wake-up message never arrives, and logs that it did. Calls that arrive while Word is closing fail instead of waiting forever.
- **Work posted to Word's thread must wait until Word is idle.** A posted message can be dispatched while Word is still handling a selection change. `InsertXML` then fails with "XML markup cannot be inserted in the specified location" (intermittently). `UiThread` posts to a message-only window, which starts a 0 ms window timer; the queue runs on WM_TIMER, which only comes when Word's queue is empty.
- **Windows 11 "Eco" quality of service.** A hidden Word, or one in the background, gets its threads run on efficiency cores at low clock: 4–5× slower on this hybrid Core 7 240H. The soil example took 42 s instead of 8 s, and drawing its results took 10–13 s instead of 1.5 s. Work the user waits for (calculation, preparing and drawing) opts out with `SetThreadInformation(ThreadPowerThrottling)` (`FullSpeedScope`), and the setting is restored afterwards.
- **Never show a message box in a hidden Word**: nobody can close it, and Word hangs. `ShowMessage` only logs when Word is invisible.
- **Opening Word during an automation run** can put the user's window into the test's hidden Word. The end-to-end test detects that (Word became visible) and leaves that Word running instead of quitting it, because quitting would discard the user's documents.
- **PowerShell 5.1 `Get-Content`/`Set-Content` re-encode UTF-8 files without a BOM as ANSI** and corrupt `′`, NBSP and so on. Edit source files with an editor, or with `[IO.File]` and an explicit UTF-8 encoding.

## Status

- Engine: 133 unit tests passing, including Calcpad's soil stress example (as one block, and converted to Word text).
- Engine tests: 136 since 0.6.0.
- Installed with the installer, version 0.6.0. The end-to-end test has 117 checks and passes against the installed build. There is one intermittent hang; see the open issue under Roadmap.
- Add-in: installed for the current user. The end-to-end test (`tools/Test-InWord.ps1`) passes, covering:
  - loading;
  - ribbon icons and context-menu IDs;
  - Evaluate, and `{= }` pick-up;
  - all switches;
  - the error count and Next Error;
  - Edit Source, then **evaluate on leaving**;
  - fixing an error by leaving it;
  - headers, footers and text boxes;
  - settings and style switching;
  - copy and paste;
  - save and reopen;
  - **blocks**:
    - fenced blocks with `#if`, a loop and a plot;
    - inline expressions using block variables;
    - editing a block, then evaluating on leaving;
    - an empty block from the button;
    - errors inside blocks;
    - a script-like selection becoming a block, with an unclosed `#if` reported;
    - blocks surviving save and reopen;
  - the soil example **calculating in the background**: Evaluate returns at once, typing meanwhile works and is kept, and the results arrive about 10 s later;
  - **Cancel** during the soil example's background calculation: it stops, nothing is drawn, and Recalculate then finishes it;
  - **math fonts**: the starters are offered in alphabetical order (Fira Math, not Minion Math), an unknown font is refused, and each free font is installed if needed and then used by every run of the equations;
  - single-step undo;
  - clean Word shutdown.
- The user ran it interactively on 2026-11-17: the ribbon works.
- **Not verified automatically** (a hidden Word has no focus or ribbon):
  - the keyboard shortcuts;
  - double-click to edit;
  - the right-click menu items;
  - the Errors menu;
  - the settings refreshing when you switch documents.
  - The test does check the add-in's answers behind these: what's under the cursor, the error list, and the menu IDs.

## Roadmap

Done:
- ribbon check;
- keyboard shortcuts;
- error list and navigation (a ribbon menu rather than a task pane, because a .NET Core add-in can't host a WinForms control in an Office task pane);
- headers, footers, text boxes and notes;
- right-click Edit Source;
- **multi-line blocks with the whole-document worksheet** (approved 2026-11-17).

Also done:
- **Script to Word text** (`ScriptConverter`): a pasted script's headings and text become native Word paragraphs, and only what needs Calcpad stays in blocks.
- **Unchanged blocks aren't redrawn**, using the input hash.
- **Per-user installer** (`CalcpadWord-Setup-<version>.exe`).
- **Background calculation** (0.4.0). Reading the document and drawing happen on Word's thread, and Calcpad runs on a worker thread.
  - If it finishes within 250 ms, the results are drawn at once (one Undo step, as before).
  - Otherwise Word stays usable. When the results come, they are drawn only if the Calcpad items are unchanged (IDs, tags, edited source); if not, it calculates again.
  - A newer calculation cancels the running one (`ExpressionParser.Cancel`), and calculations are chained, one at a time.
  - Calcpad is warmed up at start-up.
- **Math font chooser** (0.4.0, revised in 0.5.0).
  - The chosen font is stored as Word's `OMathFontName`, and every math run gets it in `w:rFonts`. Word would otherwise show runs without a font in Cambria Math.
  - Starters: the ten the user listed, except that Minion Math (commercial) was swapped for **Fira Math** in 0.5.0, at the user's request.
  - **Autodiscovery** (0.5.0): `InstalledMathFonts` reads the table directory of every registered font file (HKLM and HKCU `Fonts`) and offers each family with a MATH table. Only fonts with a MATH table have their name table read.
    - Each file is read once per session; 545 files took 0.19 s.
    - It runs in the background at start-up.
  - The list is alphabetical by label.
  - Free fonts are downloaded and installed per user (`%LOCALAPPDATA%\Microsoft\Windows\Fonts` plus HKCU, `AddFontResource`, `WM_FONTCHANGE`) after a check of the name table and the MATH table. Downloads can be cancelled.
  - On this machine all nine free fonts are installed.
- **Cancel** (0.5.0): a ribbon button, enabled only while something can be cancelled, plus Esc, taken from Word only then.
  - It cancels the running calculation (Calcpad stops within about 0.2 s, even inside the soil example's maps) and any font download.
  - Results drawn before keep their old values.

- **Out-of-date results are greyed** (0.6.0) while a background calculation runs, and after Cancel.
  - What is greyed: results already drawn, from the first new or edited item on (evaluation order), or all of them for Recalculate or a settings change.
  - The mark: light grey character shading (`wdColorGray10`), pictures in watermark mode (`msoPictureWatermark`), and the title prefixed "Out of date: ". The title keeps the mark in saved documents.
  - Unchanged results are not redrawn: only the mark is undone. That restores the document exactly, apart from Word's own bookkeeping (`wp14:editId`/`anchorId`, and `wp:effectExtent` on some pictures, 0.75 pt).
  - Fading the text (`Font.Fill.Transparency`) was tried and rejected: undoing it merges runs and loses Calcpad's per-run colours.
- **Scientific notation** (0.6.0): Calcpad's Word output writes the × of 5.7×10⁻⁹ as a bold, small, superscript run, and the exponent's minus as a hyphen in normal text. `WordMarkup.ApplyStyle` turns both into ordinary math characters (×, U+2212). `RenderVersion` went to 2, so blocks are redrawn once.

**Open issue: an intermittent hang in the hidden test Word** (seen 2026-11-17, 20:33–21:20).
- What happens: 4 of about 20 end-to-end runs hung at step 9, right after the first background result was drawn (a block with a plot). The test's next call into Word never returned.
- What is known:
  - Word's UI thread was *Running* (wait-chain traversal), not waiting on a lock, a COM call or another process.
  - No add-in code was on any stack (`dotnet-stack`), and the add-in had finished its work.
  - No wake-up message was lost: the safety timer never fired.
  - The ribbon was not involved: a hidden Word only queries it at start-up.
  - It happened with and without greying. Since then: 11 clean runs, the last 7 in a row.
- It has never been seen in a visible Word.
- Next step if it returns: sample the test Word's UI thread instruction pointer to see which module it spins in. A sampler script is written; suspending threads needs the user's approval.

Next:
1. **Code signing** for the installer, so SmartScreen doesn't warn on other machines. This needs a certificate.
2. **Configurable shortcuts.** For example, read them from `%LOCALAPPDATA%\CalcpadWord\settings.json`.
3. **Plot size.** A document default for `PlotWidth`/`PlotHeight` might help.
4. **`#read` / `#write` paths.** Make them relative to the document's folder, like `#include`.

## Script to Word text (2026-11-17, at the user's request)

`ScriptConverter.Convert` follows Calcpad's tokenizer (`'`/`"` comments, doubled quotes, `"` = heading). It parses each line's comment HTML with HtmlAgilityPack, with the expressions as placeholders:
- headings → outline levels, then Word's built-in Heading styles;
- `<strong>`/`<em>`/`<u>`/`<sub>`/`<sup>` → run formatting;
- `<hr/>` → a bottom-bordered paragraph;
- `<!-- 'expr' -->` → `#hide` expressions in a hidden paragraph (hidden paragraph mark).

Lines that stay in a block:
- `$Plot`/`$Map`;
- structural HTML (tables, lists, divs, images), up to the closing tag;
- `#if`/loops/`#def`, until balanced;
- `#include`, `#read`, `#const`.

Neighbouring block lines merge into one block, which starts with keyword lines for the active display state.

Display keywords become per-expression switches (`#round 4` → `#round 4 x`). Keywords inline expressions can't reproduce keep the script as one block: `#md`, `#format`, `#split`, `#phasor`, `#novar`, `#input`, `#pause`, and `#round <expression>`.

Gotchas:
- **InsertXML over whole paragraphs** drops the inserted content's last paragraph mark and joins that paragraph to the following text. The add-in appends an empty `<w:p/>` to absorb this.
- **Tables block InsertXML.** Word refuses to replace a range that contains a whole table (COMException 0x800A178C, "The range cannot be deleted"). This hit Edit Source, redrawing and Block to Text on the soil example's table block (fixed in 0.3.1). Every replacement now deletes the tables in the range first (`DeleteTables`). The end-to-end test covers all three cases.
- **Tag attributes** need `&#10;` for line breaks, or XML turns them into spaces.

## Input fields (0.8.0, 2026-11-18, at the user's request: "use document control or whatever Word provides for form input; the template should be reusable")

- **What it is:** a third kind of worksheet item, `ControlKind.Input`. It is a plain-text, combo-box or drop-down content control tagged `calcpad-input:{name}[:{unit}]` (`InputField`). The worksheet adds `#hide` / `name = value` / `#show` where the field stands, and hashes those lines, so blocks after a changed input are redrawn.
- **Kept apart from expressions and blocks:** `ControlTag.IsCalcpad` doesn't match input fields, so Edit Source, greying, Enter handling and the right-click menus never see them.
  - `WorksheetControls` collects both kinds in reading order.
  - The snapshot keeps the Calcpad controls as before, plus `ItemOf` (control → worksheet item) and the input fields.
  - `Draw` maps results through `ItemOf` and shades input fields that fail.
- **Values:**
  - A number is accepted; a field with a unit also accepts a number with a unit.
  - Commas are refused with a hint, so `1,200` can't become 1.2.
  - A drop-down uses the chosen item's value, or its text if it has none.
- **When it recalculates:** leaving a field whose value differs from the last calculation's. Values are remembered by control ID for the session.
  - Greying starts at the first changed field, not at the top.
  - `StillTheSame` also compares the fields, so editing one during a background calculation restarts it.
- **From scripts:** `ScriptConverter` turns `name = ? {default}` (with an optional unit after it) into `InputRun`: the name in italic with its subscript, " = ", then the field.
  - Calcpad's older saved values (after a `\v`, tab-separated) fill fields without a `{default}`. They are only used if their count matches.
  - Inputs inside blocks become `name = default`.
  - Hidden inputs become `#hide name = default`.
- **Design checks:** PASS and FAIL in block output use `class="ok"` and an inline red colour. Calcpad's `class="err"` counts as a calculation error (`CalcpadHtml.Errors`), and a failing design is not one.
- **First template, the Dokimos RC beam to EC2:** `Desktop\Alemi Herbert\Dokimos\Template source\beam`.
  - The script is `RC beam design (EN 1992-1-1).cpd`; it also runs in Calcpad desktop.
  - `Make-BeamTemplate.ps1` builds the template from Dokimos v3 and `Test-BeamTemplate.ps1` uses it as a designer would. Their scope and order of checks follow the Tedds EU library's RC beam calc.
  - A copy of the script is the fixture for `BeamTemplateTests`. That copy is checked against a hand calculation: w_Ed, M_Ed, V_Ed, d, f_cd, V_Rd,c and L/d.
- **Tested 2026-11-18:**
  - 168 unit tests.
  - `Test-BeamTemplate` (10 checks): typing and leaving a field recalculates; a failing beam says FAIL with 0 calculation errors; C40/50 sets f_ck = 40; a typo is flagged and then cleared.
  - The full end-to-end test passed on 0.8.0. A first run failed once in step 4: its `WaitFor` treated a control being redrawn as missing. `WaitFor` now retries instead.

## Installer

- `src/CalcpadWord.Setup` is a single-file `WinExe` with the zipped add-in embedded. `tools/Build-Installer.ps1` builds it.
- It installs per user, into `%LOCALAPPDATA%\Programs\CalcpadWord\app-<version>-<time>`, with HKCU registration plus an entry in Settings → Apps.
- **Shareable as one program** (0.7.0, 2026-11-18, at the user's request: "teammates don't have Calcpad").
  - **Calcpad:** its engine was already compiled in. Only `Calcpad.Core` and `Calcpad.OpenXml` ship, with their libraries. `Build-Installer.ps1` now drops Calcpad's Bulgarian and Chinese resources. They are project satellites, which `SatelliteResourceLanguages` doesn't filter. SkiaSharp (11 MB) stays: Calcpad draws plots with it, and so does the OpenXml writer.
  - **Setup:** self-contained, trimmed and compressed (10.6 MB without the payload), so it runs without .NET.
  - **Fonts:** the nine free math fonts and Libre Franklin (18 styles; the Dokimos template's font) are in `installer/fonts`. Setup installs them per user, skipping any installed for all users. Uninstall leaves them. The add-in reinstalls a removed math font from its bundled copy before trying a download.
  - **.NET runtime:** a private copy goes in `runtime\` (`host\fxr\<v>` and `shared\Microsoft.NETCore.App\<v>`, about 76 MB unpacked). Native AOT was ruled out: Word drives the add-in through classic COM (`IDispatch` ribbon callbacks, the Word object model, events), which AOT doesn't support, and it needs the MSVC linker.
  - **Why a loader:** .NET's COM host never uses an app-local runtime ("getting a delegate does not support self-contained"). It searches only `DOTNET_ROOT`, the HKLM install location and Program Files, but it reuses a `hostfxr.dll` that is already loaded in the process, with that runtime's root.
    - So Word loads `CalcpadWord.Loader.dll` (C, kernel32 only), which loads the bundled hostfxr, then forwards `DllGetClassObject` to the COM host.
    - Verified 2026-11-18: in a test process that preloaded a private hostfxr copy, the add-in's COM activation loaded `hostpolicy` and `coreclr` from that copy ("Found previously loaded library hostfxr.dll").
  - **Rejected alternatives:**
    - A per-user `DOTNET_ROOT` would redirect every .NET app the user runs.
    - Running Microsoft's runtime installer needs administrator rights.
    - A .NET Framework shim would load a second CLR into Word.
  - **Zig:** the loader is built with Zig's C compiler (`tools/Build-Loader.ps1`), since this machine has no C toolchain. The user approved the download. It is built with `-nostdlib`, which also drops Zig's libc include path, so the Windows headers are passed with `-isystem`. The result is a 7 KB DLL, kept beside its source.
  - **Size:** Setup is 56 MB. The payload is 104 MB unpacked: the runtime 77 MB, SkiaSharp 11 MB, Open XML SDK 6.5 MB and the fonts 7 MB.
  - **Tested 2026-11-18:**
    - The loader on its own: `DllGetClassObject` returned the class factory, an unknown class gave `CLASS_E_CLASSNOTAVAILABLE`, and the runtime came from `runtime\`.
    - The installed 0.7.0 in a hidden Word:
      - the add-in was connected and `a = 2 + 3` evaluated to 5;
      - all ten math fonts were installed;
      - `hostfxr`, `hostpolicy` and `coreclr` were loaded from `app-0.7.0-…\runtime\`.
    - Setup installed all 27 fonts. The full end-to-end test hasn't been rerun on 0.7.0.
- On uninstall it removes the folder after it exits, through a hidden `cmd`.
- It refuses 32-bit Office: it checks the Click-to-Run `Platform` value, or the MSI `InstallRoot` in the 64- and 32-bit registry views.
- No WiX or MSI, which avoids the extra build tools and WiX v6's licence terms.
- Tested 2026-11-17: install, the full end-to-end test on the installed copy, uninstall (everything removed), and reinstall.
