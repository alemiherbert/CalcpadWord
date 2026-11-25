# Templates

Word templates in [`templates/`](../templates/): a design calculation layout, and five complete calculations built on it. Double-click one to start a document from it.

| Template | What it calculates |
|---|---|
| Design Calculation | Empty: cover, contents, scope, design basis, inputs, calculations, results |
| RC Beam Design (EC2) | Simply supported RC beam to EN 1992-1-1: bending, shear, cracking, deflection, fire |
| Psychrometric Chart | Moist air from two known properties, with a chart (ASHRAE) |
| Duct Sizing | Round and rectangular ducts by friction rate, velocity or diameter (ASHRAE) |
| Soil Stresses and Settlement | Boussinesq stresses and settlement under a rectangular load |
| Seepage by Finite Differences | Seepage under a sheet pile wall: flow net, exit gradient, heave (EN 1997-1) |

- Fill in the cover once; the sheet headers repeat it.
- Type the inputs in the blue fields. The calculation is locked; Unlock (Calcpad tab) to change it.
- The top-left cell of the sheet header is left for your logo.
- Seepage needs add-in 0.1.2 or later; the others 0.1.1.

To build or change them: [Template source](../templates/Template%20source/README.md).

[Home](../README.md)
