# Calculating

- The document is one worksheet, top to bottom: variables, functions and macros are shared.
- Results are read-only. Double-click to edit; clicking outside recalculates from there on.
- Errors show as red source, with the message in the control's title and the status bar.
- One Undo reverts a whole Convert or Calculate.
- A copied result keeps its source.

## Automatic and manual

**Calculate automatically** is on by default: leaving a changed input or expression recalculates.

Off, it works like Excel's manual mode:
- a changed input waits, and the button reads *Calculate (out of date)*;
- an expression you write or edit is calculated on its own; what follows waits for Calculate.

## Long calculations

Past a quarter of a second, Calcpad runs in the background and Word stays usable.
- Results that may change are greyed until they arrive.
- Esc or Cancel stops it. Greyed results stay grey until the next Calculate.

[Home](../README.md)
