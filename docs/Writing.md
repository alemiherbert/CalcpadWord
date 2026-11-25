# Writing calculations

## Expressions

Type a formula, select it (or put the cursor in it) and press **Convert to Calcpad** (Alt+Shift+Enter). Each paragraph becomes one expression.

In a sentence, write `{= expression }`. It is picked up on the next Calculate.

```
L = 6m
q = 12kN/m
M = q*L^2/8
The span is {=#val L} and the midspan moment is {=#val M|kNm}.
```

To change a result, double-click it, edit, and click outside.

## Switches

In front of an inline expression; they affect only that one.

| Switch | Effect |
|---|---|
| `#equ` | Equation, substitution and result (default) |
| `#val` | Result only, with units |
| `#noc` | Expression, not calculated |
| `#hide` | Calculated, not shown |
| `#nosub` / `#varsub` | Substitution off / on |
| `#round n` | n decimals (0–15) |
| `#deg` / `#rad` / `#gra` | Angle units |

## Blocks

For multi-line Calcpad: `#if`, loops, `#def`, `#include`, `$Plot`, `$Map`, HTML tables.
- **Calcpad Block** button: on selected paragraphs, or on an empty line to type into.
- Or fences: a line `{=`, the script, a line `}`.

Rules:
- Close every `#if`, loop and `#def` you open.
- `#hide`, `#round` and the like inside a block apply to that block only.
- A block whose inputs haven't changed isn't redrawn, so a slow block doesn't slow edits below it.
- Tables with `class="bordered"` use Grid Table 1 Light, full width. Plots fill the text width. A picture alone on a line is centred; give an SVG `alt="name.svg"` to keep it vector.
- PASS/FAIL lines: `class="ok"` (green) or `style="color:#C00000"`. Calcpad's `class="err"` counts as an error.

## Pasted Calcpad scripts

Paste, select, press **Convert to Calcpad**.

| Script | Word |
|---|---|
| `"Title`, `<h1>`–`<h3>` | Heading 1 |
| `<h4>`–`<h6>` | Heading 2–4 |
| `'text 'expression' text` | Paragraph with the value inline |
| `<b>`, `<i>`, `<u>`, `<sub>`, `<sup>` | That formatting |
| `<hr/>` | Rule |
| Expression alone on a line | Equation |
| `$Plot`, `$Map`, tables, lists, `#if`, loops, `#def`, `#include` | Block |
| `#round`, `#nosub`, `#val`, `#hide` … | Switches on what follows |
| `b = ? {300}` | [Input field](Input-fields.md) |
| `<!--'x = 1'-->` | Hidden |

A script using `#md`, `#format`, `#split`, `#phasor`, `#novar`, `#input` or `#pause` stays one block. **Block to Text** converts an existing block the same way.

## Operators and symbols

- Two-key operators become Calcpad's symbols: `==` ≡, `!=` ≠, `<=` ≤, `>=` ≥, `&&` ∧, `||` ∨, `^^` ⊕, `//` ÷, `%%` ⦼, `<<` ∠, `<*` ←.
- The **Symbols** grid types Calcpad's operators, unit marks and Greek letters.

[Home](../README.md)
