using System.Text;

namespace CalcpadWord.Engine;

/// <summary>What a symbol is: for its place in the grid, and for tests.</summary>
public enum SymbolKind { Operator, Mark, Letter }

/// <summary>A symbol on the Calcpad tab: the character it inserts, its name, and the keys that also type it.</summary>
public sealed record Symbol(char Character, string Name, SymbolKind Kind, string? Typed = null);

/// <summary>
/// The symbols Calcpad uses that a keyboard doesn't have, after Calcpad's own keypad: its
/// operators, the units and marks allowed in names, and Greek letters. Greek capitals that look
/// like Latin ones (A, B, E…) are left out: they would make names that look alike but differ.
/// On the ribbon they are a grid like Word's equation Symbols gallery: <see cref="VisibleRows"/>
/// rows of <see cref="Columns"/> at a time, scrolled with ▲ ▼, and ▾ for all of them.
/// </summary>
public static class SymbolPalette
{
    public const int Columns = 10;
    public const int VisibleRows = 2;

    /// <summary>A grid cell's size on the ribbon (small buttons), in pixels at 100 % scale.</summary>
    public const int SlotSize = 16;

    /// <summary>A cell's size in the drop-down with all the symbols (as Word's), in pixels at 100 % scale.</summary>
    public const int ItemSize = 24;

    /// <summary>The group's picture when the ribbon is too narrow and folds it into one button.</summary>
    public const int GroupIconSize = 32;

    private static Symbol Op(char c, string name, string typed) => new(c, name, SymbolKind.Operator, typed);
    private static Symbol Mark(char c, string name) => new(c, name, SymbolKind.Mark);
    private static Symbol Letter(char c, string name) => new(c, name, SymbolKind.Letter);

    /// <summary>
    /// Every symbol, in grid order, in rows of ten: the operators for #if and if(); ∠ with the
    /// degree sign, the other units and the marks; the Greek letters; the rarest marks last.
    /// </summary>
    public static IReadOnlyList<Symbol> All { get; } =
    [
        Op('≤', "Less than or equal to", "<="), Op('≥', "Greater than or equal to", ">="),
        Op('≡', "Equal to", "=="), Op('≠', "Not equal to", "!="),
        Op('∧', "And", "&&"), Op('∨', "Or", "||"), Op('⊕', "Exclusive or", "^^"),
        Op('⦼', "Modulo (remainder)", "%%"), Op('÷', "Division on one line", "//"),
        Op('←', "Assign to an outer variable (in a block)", "<*"),

        Op('∠', "Phasor A∠φ (complex numbers)", "<<"),
        Mark('°', "Degree (°, °C, °F)"), Mark('′', "Prime (also arcminutes)"), Mark('″', "Double prime (also arcseconds)"),
        Mark('‴', "Triple prime"), Mark('‾', "Overline"), Mark('ø', "Diameter, ø"), Mark('Ø', "Diameter, Ø"),
        Mark('∡', "Angle"), Mark('‰', "Per mille"),

        Letter('α', "alpha"), Letter('β', "beta"), Letter('γ', "gamma"), Letter('δ', "delta"), Letter('ε', "epsilon"),
        Letter('ζ', "zeta"), Letter('η', "eta"), Letter('θ', "theta"), Letter('ϑ', "theta (script)"), Letter('ι', "iota"),
        Letter('κ', "kappa"), Letter('λ', "lambda"), Letter('μ', "mu (also micro, as in μm)"), Letter('ν', "nu"),
        Letter('ξ', "xi"), Letter('π', "pi (also the constant π)"), Letter('ρ', "rho"), Letter('σ', "sigma"),
        Letter('ς', "sigma (final)"), Letter('τ', "tau"), Letter('υ', "upsilon"), Letter('φ', "phi"),
        Letter('ϕ', "phi (straight)"), Letter('χ', "chi"), Letter('ψ', "psi"), Letter('ω', "omega"),
        Letter('Γ', "Gamma"), Letter('Δ', "Delta"), Letter('Θ', "Theta"), Letter('Λ', "Lambda"), Letter('Ξ', "Xi"),
        Letter('Π', "Pi"), Letter('Σ', "Sigma"), Letter('Φ', "Phi"), Letter('Ψ', "Psi"), Letter('Ω', "Omega (also the unit ohm)"),

        Mark('⁗', "Quadruple prime"), Mark('‱', "Per ten thousand"), Mark('℧', "Mho (siemens)"),
    ];

    /// <summary>The grid's rows, of <see cref="Columns"/> symbols (the last may be shorter).</summary>
    public static IReadOnlyList<IReadOnlyList<Symbol>> Rows { get; } =
        [.. All.Chunk(Columns).Select(row => (IReadOnlyList<Symbol>)row)];

    /// <summary>How far ▼ scrolls: the first row shown is 0 to this.</summary>
    public static int MaxTopRow => Rows.Count - VisibleRows;

    /// <summary>The symbol in a row and column of the grid, or null for an empty cell.</summary>
    public static Symbol? At(int row, int column) =>
        row >= 0 && row < Rows.Count && column >= 0 && column < Rows[row].Count ? Rows[row][column] : null;

    /// <summary>A symbol's tooltip: the symbol and its name.</summary>
    public static string ScreenTip(Symbol s) => $"{s.Character}   {s.Name}";

    /// <summary>The tooltip's second line: the keys that also type it, if any.</summary>
    public static string SuperTip(Symbol s) => s.Typed is null ? "" : $"Or type {s.Typed}";

    // ---- Ribbon ids -------------------------------------------------------------------------

    private const string ItemPrefix = "CalcpadSym_";
    private const string IconPrefix = "CalcpadSymIcon_";
    private const string SlotPrefix = "CalcpadSlot_";
    public const string UpImage = "CalcpadSymNavUp";
    public const string DownImage = "CalcpadSymNavDown";
    public const string MoreImage = "CalcpadSymNavMore";

    /// <summary>The ribbon id of a symbol's item in the drop-down, from which <see cref="FromItemId"/> gets the character back.</summary>
    public static string ItemId(char c) => ItemPrefix + ((int)c).ToString("X4");

    public static char? FromItemId(string? id) => Parse(id, ItemPrefix);

    /// <summary>The ribbon id of a grid cell (row on screen, column).</summary>
    public static string SlotId(int row, int column) => $"{SlotPrefix}{row}_{column}";

    public static (int Row, int Column)? FromSlotId(string? id)
    {
        if (id is null || !id.StartsWith(SlotPrefix, StringComparison.Ordinal))
            return null;
        var parts = id[SlotPrefix.Length..].Split('_');
        return parts.Length == 2 && int.TryParse(parts[0], out var row) && int.TryParse(parts[1], out var column) &&
               row is >= 0 and < VisibleRows && column is >= 0 and < Columns
            ? (row, column)
            : null;
    }

    public static IEnumerable<string> SlotIds() =>
        Enumerable.Range(0, VisibleRows).SelectMany(r => Enumerable.Range(0, Columns).Select(c => SlotId(r, c)));

    /// <summary>
    /// The picture a ribbon <c>image</c> name asks for (customUI's loadImage): a symbol's cell
    /// in the drop-down, or the group's icon.
    /// </summary>
    public static (char Character, bool Icon)? FromImageId(string? id) =>
        Parse(id, IconPrefix) is { } icon ? (icon, true) :
        Parse(id, ItemPrefix) is { } item ? (item, false) :
        null;

    private static char? Parse(string? id, string prefix) =>
        id is not null && id.Length > prefix.Length && id.StartsWith(prefix, StringComparison.Ordinal) &&
        int.TryParse(id.AsSpan(prefix.Length), System.Globalization.NumberStyles.AllowHexSpecifier, null, out var code) &&
        code is > 0 and <= 0xFFFF
            ? (char)code
            : null;

    /// <summary>
    /// The Symbols group's controls, as ribbon markup (customUI 2009/07): two rows of cells
    /// (small buttons, filled by callbacks as the grid scrolls), then, after a separator so that
    /// they make a column of their own, ▲, ▼ and ▾ (a drop-down gallery of every symbol, in the
    /// same rows).
    /// </summary>
    public static string RibbonXml()
    {
        var sb = new StringBuilder();
        for (var r = 0; r < VisibleRows; ++r)
        {
            sb.Append($"<box id=\"CalcpadSymbolRow{r}\" boxStyle=\"horizontal\">");
            for (var c = 0; c < Columns; ++c)
                sb.Append($"<button id=\"{SlotId(r, c)}\" showLabel=\"false\" getLabel=\"GetSlotLabel\" getImage=\"GetSlotImage\" ")
                  .Append("getScreentip=\"GetSlotScreentip\" getSupertip=\"GetSlotSupertip\" getEnabled=\"GetSlotEnabled\" onAction=\"OnSlot\" />");
            sb.Append("</box>");
        }
        sb.Append("<separator id=\"CalcpadSymbolsScrollBar\" />")
          .Append($"<button id=\"CalcpadSymbolsUp\" showLabel=\"false\" label=\"Previous row\" image=\"{UpImage}\" ")
          .Append("getEnabled=\"GetSymbolsUpEnabled\" onAction=\"OnSymbolsUp\" screentip=\"Previous row\" />")
          .Append($"<button id=\"CalcpadSymbolsDown\" showLabel=\"false\" label=\"Next row\" image=\"{DownImage}\" ")
          .Append("getEnabled=\"GetSymbolsDownEnabled\" onAction=\"OnSymbolsDown\" screentip=\"Next row\" />")
          .Append($"<gallery id=\"CalcpadSymbolsAll\" showLabel=\"false\" label=\"All Symbols\" image=\"{MoreImage}\" ")
          .Append($"columns=\"{Columns}\" itemWidth=\"{ItemSize}\" itemHeight=\"{ItemSize}\" showItemLabel=\"false\" onAction=\"OnSymbol\" ")
          .Append($"screentip=\"All symbols\" supertip=\"{Escape("Operators, units and marks, Greek letters. Operators can also be typed with two keys: >= becomes ≥ when calculated.")}\">");
        foreach (var symbol in All)
        {
            var id = ItemId(symbol.Character);
            sb.Append($"<item id=\"{id}\" label=\"{Escape(symbol.Character.ToString())}\" image=\"{id}\" ")
              .Append($"screentip=\"{Escape(ScreenTip(symbol))}\"")
              .Append(symbol.Typed is null ? "" : $" supertip=\"{Escape(SuperTip(symbol))}\"")
              .Append(" />");
        }
        sb.Append("</gallery>");
        return sb.ToString();
    }

    private static string Escape(string s) => System.Security.SecurityElement.Escape(s);
}
