# Input fields

A Word content control (plain text, combo box or drop-down) that sets a Calcpad variable. Users type a number or pick an item, and never see Calcpad.

- **Make one:** write `name = ? {default}` in a script and press Convert to Calcpad. By hand: a content control tagged `calcpad-input:name`.
- **Units:** the tag `calcpad-input:L:m` gives a bare number the unit m. Without a unit, only a plain number is accepted.
- **E notation:** `1e-5`, `2.5E-7`.
- **Drop-downs:** an item's value is the number used, its display name is what's shown (*C30/37* → 30). Values must be unique: write 0.3 and 0.30 if two items mean the same.
- **Wrong entries** (`0,85`, a typo) are shaded red and listed under Errors. The rest still calculates.
- Leaving a changed field recalculates.
- Inside a block, an input stays an ordinary assignment.

[Home](../README.md)
