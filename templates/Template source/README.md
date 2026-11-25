# Template source

Builds and tests the calculation templates in `templates/`. Needs desktop Word, Calcpad for Word and Python 3. Each Word script runs its own hidden Word and takes `-OutDir <folder>`, where it writes `steps.log` and `pid.txt`.

## Generic and Dokimos

The templates on GitHub are generic. On the author's PCs, `templates/dokimos/` (not on GitHub) holds Dokimos versions: the same files with the company's logo and name. Where that folder is, the Make scripts build the Dokimos version and its generic twin; elsewhere they build from the generic base.

`make_generic.py "<Dokimos X vN.dotx>" "<X vN.dotx>"` makes a twin: it removes the logo and the name, copies everything else byte for byte, and fails if "Dokimos" is left.

## Rules

- Every new template is a new version, `<Name> vN.dotx`. Only the latest is kept; git keeps the rest.
- No guidance lines: a calculation is for submission.
- A template doesn't say where its calculation came from.

## Making one

- The calculation is a Calcpad script (`.cpd`). `name = ? {default}` lines become input fields; `'<!--section:NAME-->` lines mark where each part goes.
- `<calc>\Make-…Template.ps1 -OutDir <dir> -Version N` builds `<calc>\out\<Name> vN.dotx`: cover and design basis filled, each section converted by the add-in, drop-downs, calculated, locked (inputs open), page pictures. `TemplateBuild.ps1` holds the shared steps.
- `<calc>\Test-…Template.ps1 -OutDir <dir> [-Template <file>]` uses it as a designer would and logs ok or FAIL.
- `Test-Template.ps1 -OutDir <dir> [-Template <file>]` checks the base: the cover's project shows in the header, no guidance, a Calcpad line converts.
- The unit tests read the `.cpd` files from here: `dotnet test tests\CalcpadWord.Engine.Tests`.
- Copy the finished version to `templates/`.

Script pitfalls:
- No apostrophes in comments (they end the comment). No variables named like units (m, g, A, N, V, P).
- Put a dash in empty table cells: Word refuses a block with an empty `<td>`.
- Keep hidden `#if` groups in one block, ending with something visible.

## The templates

| Folder | Template | Checked against |
|---|---|---|
| `beam/` | RC Beam Design (EC2): cover, fire, bending, limits, spacing, shear, cracking, deflection (EN 1992-1-1) | Tekla Tedds (`Check-WithTedds.ps1`), `BeamTemplateTests` |
| `psychrometric/` | Psychrometric Chart: moist air from two known properties, with a chart (ASHRAE Fundamentals ch. 1) | ASHRAE's tables, `PsychrometricTests` |
| `duct/` | Duct Sizing: round and rectangular ducts by friction rate, velocity or diameter (ASHRAE ch. 21) | McQuay DuctSizer, `DuctTests` |
| `soil/` | Soil Stresses and Settlement: Boussinesq under a rectangular load varying across its width | Calcpad's example, classical results, `SoilTests` |
| `seepage/` | Seepage by Finite Differences: sheet pile wall, flow net, exit gradient, heave (EN 1997-1) | Exact solution, Craig's example, `SeepageTests` |

Design notes and version history are in `PLAN.md`.

## Design

- A4, 0.6 in margins; Libre Franklin, SemiBold headings.
- A bottom-aligned cover: title, details and document control at the foot; a faint corner pattern inside the margins.
- Contents numbered i, ii; body 1, 2 ("Sheet x of y" counts the body). No footers.
- Cover fields are stored once, in a custom XML part (`urn:calcpad-word:calculation`), and shown again in the sheet headers.
