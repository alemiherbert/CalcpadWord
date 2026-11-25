# Tips and limits

## Word's AutoCorrect

- Turn off File → Options → Proofing → AutoCorrect Options → AutoFormat As You Type → *\*Bold\* and \_italic\_ with real formatting*. It deletes `*` and `_` from formulas.
- *Capitalize first letter of sentences* turns `f(x)` into `F(x)` at the start of a line, and Calcpad is case-sensitive. Press Ctrl+Z straight after, or turn it off.
- En dashes, smart quotes, ½ and non-breaking spaces are fixed automatically.

## Equations

Calcpad's output is corrected to read as Word math: real minus signs, upright subscripts (M_Ed), regular function names, and × 10ⁿ for small and large numbers.

## Limits

- Inline expressions can't hold quotes or plots: use a block.
- Calcpad stops after 40 errors; the rest is marked *Not calculated*.
- Comments in Word's review pane aren't scanned.

## Logs

Add-in: `%LOCALAPPDATA%\CalcpadWord\addin.log`. Setup: `%TEMP%\CalcpadWord-Setup.log`.

[Home](../README.md)
