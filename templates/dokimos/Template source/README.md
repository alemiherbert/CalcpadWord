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

`Dokimos RC Beam Design (EC2) vN.dotx` is the Dokimos design calculation template (v3) with a complete calculation in it. It designs a simply supported rectangular RC beam to EN 1992-1-1:
- bending;
- reinforcement limits and bar spacing;
- shear;
- crack width;
- span to depth ratio.

Each check ends in PASS or FAIL, and a summary table follows. It needs Calcpad for Word 0.8.0 or later, which reads its input fields.

**Using it:** fill in the blue fields under Design inputs, by typing a value or picking one from a list. Leaving a field recalculates.

**How it is made:**
- `RC beam design (EN 1992-1-1).cpd` is the calculation, an ordinary Calcpad script; it also runs in Calcpad desktop.
  - `name = ? {default}` lines become the input fields.
  - `'<!--section:NAME-->` lines mark where each part goes: materials, geometry, loads, calculations, results.
- `Make-BeamTemplate.ps1 -OutDir <dir> -Version N` builds `out\Dokimos RC Beam Design (EC2) vN.dotx`:
  1. it opens a new document from v3 in a hidden Word;
  2. it fills the cover, scope, codes and assumptions;
  3. it puts each part of the script under its heading and lets Calcpad for Word convert it;
  4. it turns some inputs into drop-down lists: concrete class, exposure class, category of use, bar sizes;
  5. it renders the pages to `out\vN-pageM.png`.
- `Test-BeamTemplate.ps1 -OutDir <dir> -Version N` uses the template as a designer would and logs ok or FAIL for each check. The add-in's unit tests (`BeamTemplateTests`) also calculate the script and check it against a hand calculation.

**Changing it:** edit the `.cpd`, then run Make-BeamTemplate with the next version number, and Test-BeamTemplate. Copy the new `.dotx` to the folder above; never overwrite an earlier version. The add-in's unit tests (`BeamTemplateTests`) read the `.cpd` from here, so run them too: `dotnet test tests\CalcpadWord.Engine.Tests`.

**A new calculation template:** do the same with another script. Make-BeamTemplate shows each step:
1. Write the Calcpad script with `?` inputs and section markers.
2. Copy Make-BeamTemplate under a new name and change the scope, codes and assumptions text and the drop-down lists.
3. Build it, look at the pages and test it.

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
