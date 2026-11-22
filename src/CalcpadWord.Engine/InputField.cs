using System.Text.RegularExpressions;

namespace CalcpadWord.Engine;

/// <summary>
/// A Word form field that sets a Calcpad variable: a plain-text, combo-box or drop-down list
/// content control tagged <c>calcpad-input:{name}</c> (or <c>calcpad-input:{name}:{unit}</c>).
/// Whoever fills in the document types a number, or picks an item whose value is the number,
/// and the worksheet reads it as <c>name = number</c> where the field stands. Calcpad's own
/// input syntax, <c>name = ? {default}</c>, becomes such a field when a script is converted.
/// </summary>
/// <param name="Unit">Appended to a bare number (e.g. "mm"). Empty for unitless worksheets: then only a number is accepted.</param>
public sealed partial record InputField(string Name, string Unit = "")
{
    public const string Prefix = "calcpad-input:";

    /// <summary>A Calcpad variable name: a letter or _, then letters, digits, _, primes, commas and points (A_s,req).</summary>
    [GeneratedRegex(@"^[\p{L}_][\p{L}\p{N}_′″‴⁗,.]*$")]
    internal static partial Regex VariableName();

    /// <summary>A number, also in E notation as typed in Excel ("1e-5", "2.5E-7"), which Calcpad doesn't read.</summary>
    [GeneratedRegex(@"^(?<number>-?(\d+\.?\d*|\.\d+))([eE](?<exponent>[-+]?\d+))?$")]
    private static partial Regex Number();

    /// <summary>A number with a unit, accepted where the field has a unit: "0.3 m", "25kN/m", "1e-5 m/s".</summary>
    [GeneratedRegex(@"^(?<number>-?(\d+\.?\d*|\.\d+))([eE](?<exponent>[-+]?\d+))?\s*(?<unit>[\p{L}°][\p{L}\p{N}°/·*^.]*)$")]
    private static partial Regex NumberWithUnit();

    /// <summary>The number as Calcpad writes it, with the unit after it: "1e-5" with m/s is 1*10^(-5)*1m/s.</summary>
    private static string Quantity(Match m, string unit)
    {
        var number = m.Groups["number"].Value;
        if (!m.Groups["exponent"].Success)
            return number + unit;
        var exponent = int.Parse(m.Groups["exponent"].Value, System.Globalization.CultureInfo.InvariantCulture);
        return $"{number}*10^({exponent})" + (unit.Length > 0 ? "*1" + unit : string.Empty);
    }

    public static bool IsInput(string? tag) => tag is not null && tag.StartsWith(Prefix, StringComparison.Ordinal);

    public static InputField? Parse(string? tag)
    {
        if (!IsInput(tag))
            return null;

        var parts = tag![Prefix.Length..].Split(':', 2);
        var name = parts[0].Trim();
        return VariableName().IsMatch(name) ? new InputField(name, parts.Length > 1 ? parts[1].Trim() : string.Empty) : null;
    }

    public override string ToString() => Prefix + Name + (Unit.Length == 0 ? string.Empty : ":" + Unit);

    /// <summary>The Calcpad line for a value from the field. Throws <see cref="FormatException"/> with a message for the user.</summary>
    public string Assignment(string? value)
    {
        var text = SourceNormalizer.Normalize(value).TrimStart('+');
        if (text.Length == 0 || text == "?")
            throw new FormatException($"Enter a value for {Name}.");
        if (Number().Match(text) is { Success: true } n)
            return $"{Name} = {Quantity(n, Unit)}";
        if (Unit.Length > 0 && NumberWithUnit().Match(text) is { Success: true } m)
            return $"{Name} = {Quantity(m, m.Groups["unit"].Value)}";
        if (text.Contains(',') && Number().IsMatch(text.Replace(",", "")))
            throw new FormatException($"{Name}: write the number without commas: a point for decimals (0.85), nothing between thousands (1200).");
        throw new FormatException(Unit.Length == 0
            ? $"{Name}: enter a number, not \"{text}\"."
            : $"{Name}: enter a number (in {Unit}), not \"{text}\".");
    }
}
