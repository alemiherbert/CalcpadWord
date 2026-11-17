namespace CalcpadWord.Engine;

public enum ControlKind
{
    /// <summary>One expression inside a sentence or paragraph.</summary>
    Inline,
    /// <summary>A multi-line Calcpad script occupying whole paragraphs.</summary>
    Block,
    /// <summary>A Word form field that sets a variable (see <see cref="InputField"/>). It has no output.</summary>
    Input,
}

public enum ControlState
{
    /// <summary>Shows the rendered output; the source is kept in the tag.</summary>
    Result,
    /// <summary>Shows the source as editable text (Edit Source); read back on the next evaluation.</summary>
    Editing,
    /// <summary>Shows the source as editable red text because it could not run; read back on the next evaluation.</summary>
    Error,
}

/// <summary>
/// What a Calcpad content control remembers, stored in its Tag:
/// <c>calcpad:{r|e|x}:{fingerprint|-}:{source}</c> for inline expressions and
/// <c>calcpad-block:…</c> for blocks. Word keeps tags of any Unicode, line breaks included,
/// up to at least 140,000 characters, so each control is self-contained: copy, paste and undo
/// carry the source along.
/// </summary>
public sealed record ControlTag(ControlState State, string Fingerprint, string Source, ControlKind Kind = ControlKind.Inline)
{
    public const string Prefix = "calcpad:";
    public const string BlockPrefix = "calcpad-block:";

    public static bool IsCalcpad(string? tag) =>
        tag is not null &&
        (tag.StartsWith(Prefix, StringComparison.Ordinal) || tag.StartsWith(BlockPrefix, StringComparison.Ordinal));

    /// <summary>The control shows source text (being edited, or failed), which is read back when evaluating.</summary>
    public bool ShowsSource => State != ControlState.Result;

    public static ControlTag? Parse(string? tag)
    {
        if (tag is null)
            return null;

        ControlKind kind;
        string rest;
        if (tag.StartsWith(Prefix, StringComparison.Ordinal))
            (kind, rest) = (ControlKind.Inline, tag[Prefix.Length..]);
        else if (tag.StartsWith(BlockPrefix, StringComparison.Ordinal))
            (kind, rest) = (ControlKind.Block, tag[BlockPrefix.Length..]);
        else
            return null;

        var parts = rest.Split(':', 3);
        if (parts.Length < 3)
            return null;

        var state = parts[0] switch
        {
            "e" => ControlState.Editing,
            "x" => ControlState.Error,
            _ => ControlState.Result,
        };
        var fingerprint = parts[1] == "-" ? string.Empty : parts[1];
        return new ControlTag(state, fingerprint, parts[2], kind);
    }

    public override string ToString()
    {
        var state = State switch
        {
            ControlState.Editing => 'e',
            ControlState.Error => 'x',
            _ => 'r',
        };
        var prefix = Kind == ControlKind.Block ? BlockPrefix : Prefix;
        return $"{prefix}{state}:{(Fingerprint.Length == 0 ? "-" : Fingerprint)}:{Source}";
    }
}
