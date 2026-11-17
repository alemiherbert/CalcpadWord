using System.Globalization;

namespace CalcpadWord.Engine;

public enum DisplayMode
{
    /// <summary>Equation, substituted values and result (Calcpad's default, <c>#equ</c>).</summary>
    Equation,
    /// <summary>Result only, with units (<c>#val</c>). Meant for numbers inside sentences.</summary>
    Value,
    /// <summary>Show the expression without calculating it (<c>#noc</c>).</summary>
    NoCalculation,
    /// <summary>Calculate silently; the expression is kept as hidden text (<c>#hide</c>).</summary>
    Hidden,
}

/// <summary>
/// One inline expression plus the Calcpad-style switches written in front of it, e.g.
/// <c>#val #round 3 M|kNm</c>. Switches only affect that one expression.
/// </summary>
public sealed record InlineDirective(
    string Expression,
    DisplayMode Mode,
    bool? Substitute,
    int? Decimals,
    AngleUnits? Angles)
{
    public const string Help = "#equ, #val, #noc, #hide, #nosub, #varsub, #round n, #deg, #rad or #gra";

    /// <exception cref="FormatException">An unknown switch or a bad <c>#round</c> value.</exception>
    public static InlineDirective Parse(string source)
    {
        var mode = DisplayMode.Equation;
        bool? substitute = null;
        int? decimals = null;
        AngleUnits? angles = null;

        var rest = source.AsSpan().Trim();
        while (rest.Length > 0 && rest[0] == '#')
        {
            var word = NextWord(ref rest).ToString();
            switch (word.ToLowerInvariant())
            {
                case "#equ": mode = DisplayMode.Equation; break;
                case "#val": mode = DisplayMode.Value; break;
                case "#noc": mode = DisplayMode.NoCalculation; break;
                case "#hide": mode = DisplayMode.Hidden; break;
                case "#nosub": substitute = false; break;
                case "#varsub": substitute = true; break;
                case "#deg": angles = AngleUnits.Degrees; break;
                case "#rad": angles = AngleUnits.Radians; break;
                case "#gra": angles = AngleUnits.Gradians; break;
                case "#round":
                    var n = NextWord(ref rest);
                    if (!int.TryParse(n, NumberStyles.None, CultureInfo.InvariantCulture, out var d) || d > 15)
                        throw new FormatException("#round needs a number of decimals from 0 to 15, for example: #round 3 x/7");
                    decimals = d;
                    break;
                default:
                    throw new FormatException($"Unknown switch \"{word}\". Use {Help}.");
            }
        }
        return new InlineDirective(rest.ToString(), mode, substitute, decimals, angles);
    }

    private static ReadOnlySpan<char> NextWord(ref ReadOnlySpan<char> s)
    {
        var end = 0;
        while (end < s.Length && !char.IsWhiteSpace(s[end]))
            ++end;

        var word = s[..end];
        s = s[end..].TrimStart();
        return word;
    }
}
