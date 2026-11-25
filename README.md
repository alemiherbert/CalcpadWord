# Calcpad for Word

[Calcpad](https://codeberg.org/Proektsoft/Calcpad) calculations in desktop Word (64-bit Windows). Formulas become Word equations in your text; plots and multi-line scripts become blocks drawn by Calcpad. The document is one worksheet, calculated top to bottom.

**Download:** [Releases](https://github.com/alemiherbert/CalcpadWord/releases), `CalcpadWord-Setup-<version>.exe`.

## Quick start

1. Run Setup and start Word: a **Calcpad** tab appears.
2. Type `L = 6m`, then press **Convert to Calcpad** (Alt+Shift+Enter).
3. For a plot, put the lines between `{=` and `}`:

```
{=
f(x) = x^2 - 5*x + 3
$Plot{f(x) @ x = -1 : 6}
}
```

## Wiki

- [Install](docs/Install.md)
- [Writing calculations](docs/Writing.md): expressions, switches, blocks, pasted scripts
- [Input fields](docs/Input-fields.md): reusable templates
- [Ribbon and shortcuts](docs/Ribbon.md)
- [Calculating](docs/Calculating.md): automatic and manual, background, errors
- [Tips and limits](docs/Tips.md)
- [Development](docs/Development.md): build, test, layout
