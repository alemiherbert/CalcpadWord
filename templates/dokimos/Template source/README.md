# Dokimos design-calculation template: source

These scripts made the templates in the folder above. They need Python 3 and desktop Word.

Each Word script runs its own hidden Word: your open documents are left alone. Each one takes `-OutDir <folder>`, where it writes `steps.log` and `pid.txt` (the hidden Word's process ID, in case it has to be stopped).

## Rules

- Every new template is a new version: `Dokimos Design Calculation vN.dotx`. Never overwrite an earlier one.
- The latest version is the base. When you adjust a document by hand (nudge the logo or pattern), that document becomes the base. Change it; don't regenerate it.

## How each version was made

| Version | Made by |
|---|---|
| v1 | `build_template.py` writes `out\dokimos-template.docx` (and `out\cover-pattern.svg`) from the styles in `unpacked-1\` (the original document, unpacked). `Finish-Template.ps1` adds the logo (`assets\logo.svg`, from dokimos.net) and the cover pattern, updates the contents, and saves the `.dotx`. |
| v2 | Your hand-adjusted `user\calculation-title.docx`, with the cover pattern fixed. `Make-Fallback.ps1` has Word make `user\fallback.docx`, which holds the corrected pattern and its PNG fallback. `swap_pattern.py user\calculation-title.docx user\fallback.docx user\calculation-title-fixed.docx` replaces only the pattern's pictures, so positions are kept. |
| v3 | `make_v3.py user\calculation-title-fixed.docx user\v3.docx` splits the document into three sections: a cover with no footer; contents numbered i, ii; a body numbered 1, 2 ("Sheet x of y" counts body pages only). It removes the footers. `Version-Template.ps1 -OutDir <dir> -In user\v3.docx -Out "user\Dokimos Design Calculation v3.dotx" -Name v3` updates the contents and fields, renders `user\v3-pageN.png`, and saves the `.dotx`. |

## Making v4

1. Start from v3: either a document made from it with your changes, or `user\v3.docx`.
2. Make the change. Edit the XML with a small script like `make_v3.py`, so that everything else stays byte for byte.
3. Run `Version-Template.ps1 -OutDir <dir> -In <your .docx> -Out "user\Dokimos Design Calculation v4.dotx" -Name v4`.
4. Look at the page pictures, then copy the `.dotx` to the folder above.

## Calculation templates (beam/)

`Dokimos RC Beam Design (EC2) vN.dotx` is the Dokimos design calculation template (v3) with a complete calculation in it. It designs a simply supported rectangular RC beam to EN 1992-1-1. v2 checks:
- cover, from the exposure class (4.4.1);
- fire resistance (EN 1992-1-2, Table 5.5);
- bending by K and K′ (6.1);
- reinforcement limits and the spacing of the top and bottom bars;
- shear: struts, links, link spacing;
- crack control: minimum steel (7.1) and crack width (7.3.4);
- deflection: span to depth ratio (7.4.2) and calculated deflection (7.4.3), with creep and shrinkage from Annex B.

A section sketch follows the geometry. Each check ends in a "PASS - …" or "FAIL - …" line, and a summary table follows. It needs Calcpad for Word 0.8.0 or later, which reads its input fields.

| Version | What changed |
|---|---|
| v1 | Bending, limits, spacing, shear, crack width, span to depth ratio. |
| v2 | Modelled on Tekla Tedds' "RC beam design (EN1992)" and Calcpad's EC2 examples: cover, fire, K and K′, top bar spacing, minimum steel for cracking, calculated deflection, the sketch, reactions. Tables use Grid Table 1 Light; plots fill the width. The exposure class sets the cover and w_max. Calculates on Recalculate only (automatic calculation off), so input is fast. |

**Checked against Tedds:** `Check-WithTedds.ps1 -OutDir <dir> [-Recommended]` runs Tedds' own calculation for the default section through its API (Tedds must be installed; it saves Tedds' output as text). The numbers agree. They are in `BeamTemplateTests` (`Agrees_with_Tedds_…`).

**Using it:** fill in the blue fields under Design inputs, by typing a value or picking one from a list, then click Recalculate (Ctrl+Alt+F9). Until then the button reads *Recalculate (out of date)*. To calculate whenever a field is left, tick *Calculate automatically* on the Calcpad tab.

## Calculation templates (psychrometric/)

`Dokimos Psychrometric Chart vN.dotx` gives the properties of moist air from any two known properties at the site's altitude, by ASHRAE Fundamentals chapter 1, with a psychrometric chart: dry and wet bulb, dew point, relative humidity, humidity ratio, enthalpy, specific volume, density, degree of saturation.

| Version | What changed |
|---|---|
| v1 | Built on the first PC by an earlier script (2026-11-18), with drop-downs for the two properties and a results table. The working was about 20 hidden paragraphs; Calculations showed one sentence. |
| v2 | Built from `Psychrometric chart.cpd`. The working is hidden; ASHRAE's relations follow as Calcpad equations (formula = result), with their constants in full. Each known property is one line: "[property] = [value]". The chart has a small legend (a coloured line and its words, one per line) and ends the calculation: no results table, which repeated the working. Red messages for pairs that can't fix the air. Calculates on Recalculate only. |

**Checked:** `PsychrometricTests` compares the standard atmosphere and the saturation pressures with ASHRAE's tables, and the air at 25 °C and 50 % with its chart. It also finds that air again from each of the 19 pairs of properties that fix it.

## Calculation templates (duct/)

`Dokimos Duct Sizing vN.dotx` replaces McQuay's DuctSizer 6.4 (Daikin). For an airflow it sizes a round duct by friction rate or velocity, or checks a given diameter (Darcy–Weisbach, Colebrook), and gives the other side of a rectangular duct for the same friction (Huebscher), rounded up to 25 mm. Air properties come from the temperature and the site's altitude; the duct material sets the roughness (ASHRAE's categories).

| Version | What changed |
|---|---|
| v1 | Built from `Duct sizing.cpd`. Calculates on Recalculate only. |

**Checked against DuctSizer:** `DuctTests` holds four of its results (three by friction rate, one by velocity) and its equivalent diameter for 500 × 200 mm. With DuctSizer's roughness the diameters agree within 0.3 % (see PLAN.md, "Duct sizing v1").

## Calculation templates (soil/)

`Dokimos Soil Stresses and Settlement vN.dotx` gives the stresses and the settlement under a rectangular surface load that varies linearly across its width, on an elastic half-space: Boussinesq integrated over the loaded area, the settlement from the vertical strain (EN 1997-1, Annex F.1). Maps of σz and τxz through the centre, depth profiles and the settlement across the width. The design inputs have sketches drawn from the inputs: a section and a plan of the loaded area (to scale), and the load across the width; the build keeps each input line on the page of its sketch.

| Version | What changed |
|---|---|
| v1 | Built from `Soil stresses and settlement.cpd`, after Calcpad's example "Soil stress and deformations" (not named in the document, at the user's request). Calculates on Recalculate only. |

**Checked:** `SoilTests` against Calcpad's example (9.96 mm), the classical settlement of a flexible rectangle on a deep half-space (12.35 mm for 2 × 3 m at 100 kPa) and Boussinesq's point load far below.

## How a calculation template is made

- **The calculation** is an ordinary Calcpad script (`.cpd`); it also runs in Calcpad desktop.
  - `name = ? {default}` lines become the input fields.
  - `'<!--section:NAME-->` lines mark where each part goes.
- **`Make-…Template.ps1 -OutDir <dir> -Version N`** builds `out\<name> vN.dotx`, with `TemplateBuild.ps1` doing the common steps:
  1. it opens a new document from v3 in a hidden Word;
  2. it fills the cover, scope, codes and assumptions, and adjusts the headings if need be;
  3. it puts each part of the script under its heading and lets Calcpad for Word convert it;
  4. it turns some inputs into drop-down lists;
  5. it recalculates, makes the empty paragraphs after the last block 1 pt (so they never add a blank page), sets the template to calculate on Recalculate only, renders the pages to `out\vN-pageM.png` and saves it.
- **`Test-…Template.ps1 -OutDir <dir> -Version N`** uses the template as a designer would and logs ok or FAIL for each check.
- **The unit tests** (`BeamTemplateTests`, `PsychrometricTests`, `DuctTests`, `SoilTests`) read the `.cpd` from here and calculate it without Word.
- **Inputs with a unit** (`a = ? {2}m`) show it after the field and are tagged with it (`calcpad-input:a:m`).
- **A template doesn't say where its calculation came from** (the user, 2026-11-20).
- **Empty table cells:** Calcpad's Word writer leaves an empty `<td></td>` without a paragraph, and Word then refuses the whole block ("XML markup cannot be inserted"). Put a dash in empty cells. (Found with psychrometric v1.)
- **Hidden working:** keep the hidden `#if` groups in one block, and end it with something visible. A `#show` between two `#if` groups is not carried into the block, and an empty block shows as a blank line. (Found with psychrometric v1.)

**Changing one:** edit the `.cpd`, run the unit tests (`dotnet test tests\CalcpadWord.Engine.Tests`), then its Make and Test scripts. Work on one new version and rebuild it until it is done. Copy it to the folder above; never overwrite a version already given out.

**A new one:** write the script with `?` inputs and section markers, then copy a Make script and change its text, headings and drop-down lists.

## Other scripts

- `Test-Template.ps1 -OutDir <dir>`: makes a new document from v3 in a hidden Word and checks three things:
  - the project typed on the cover appears in the sheet header;
  - typing over the guidance removes it;
  - a Calcpad line evaluates.

  It needs the Calcpad add-in. Change the template path in the script to test a later version.
- `check_xml.py <files>`: checks that XML files are well-formed.

## Design decisions

- **Page:** A4 with 0.6 in margins.
- **Font:** Libre Franklin (SemiBold for headings).
- **Logo:** at the top of the cover.
- **Company details:** no company name, URL or tagline anywhere.
- **Cover:** bottom-aligned, with the title, details and document control at the foot.
- **Cover pattern:** a faint geometric pattern at the top right, inside the margins so it prints on office printers. Its triangles follow the mesh's 60° lines.
- **Brand colours (dokimos.net):** ink `1C1710`, paper `F7F4EE`, blue `2E63C4`, green `2FA968`.
- **Document fields:** project, numbers, client, date and people live in a custom XML part (`urn:dokimos-net:calculation`). The fields on the cover and in the headers show the same values.
