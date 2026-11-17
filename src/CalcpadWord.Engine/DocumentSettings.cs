using System.Globalization;

namespace CalcpadWord.Engine;

public enum AngleUnits { Degrees = 0, Radians = 1, Gradians = 2 }

public enum RenderStyle
{
    /// <summary>Calcpad's own look: variable names coloured blue.</summary>
    Calcpad,
    /// <summary>Plain black Word equations.</summary>
    PlainMath,
}

/// <summary>
/// Per-document calculation settings. They are stored inside the Word document so that
/// recalculating gives the same output on any machine.
/// </summary>
public sealed record DocumentSettings
{
    private readonly int _decimals = 2;

    public int Decimals { get => _decimals; init => _decimals = Math.Clamp(value, 0, 15); }
    public bool Substitute { get; init; } = true;
    public AngleUnits Angles { get; init; } = AngleUnits.Degrees;
    public RenderStyle Style { get; init; } = RenderStyle.Calcpad;

    /// <summary>The equations' font. In Word it is the document's math font; the add-in keeps the two in step.</summary>
    public string MathFont { get; init; } = MathFonts.Default;

    public string Serialize() =>
        string.Create(CultureInfo.InvariantCulture,
            $"decimals={Decimals};substitute={(Substitute ? 1 : 0)};angles={AngleKey(Angles)};style={(Style == RenderStyle.PlainMath ? "plain" : "calcpad")};font={MathFont}");

    /// <summary>Tolerant parse: unknown keys and bad values fall back to defaults.</summary>
    public static DocumentSettings Parse(string? text)
    {
        var settings = new DocumentSettings();
        if (string.IsNullOrWhiteSpace(text))
            return settings;

        foreach (var part in text.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var eq = part.IndexOf('=');
            if (eq <= 0)
                continue;

            var key = part[..eq].Trim().ToLowerInvariant();
            var value = part[(eq + 1)..].Trim().ToLowerInvariant();
            settings = key switch
            {
                "font" when part[(eq + 1)..].Trim().Length > 0 => settings with { MathFont = part[(eq + 1)..].Trim() },
                "decimals" when int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var d)
                    => settings with { Decimals = d },
                "substitute" => settings with { Substitute = value is "1" or "true" },
                "angles" when TryParseAngles(value, out var a) => settings with { Angles = a },
                "style" => settings with { Style = value == "plain" ? RenderStyle.PlainMath : RenderStyle.Calcpad },
                _ => settings,
            };
        }
        return settings;
    }

    internal static string AngleKey(AngleUnits a) => a switch
    {
        AngleUnits.Radians => "rad",
        AngleUnits.Gradians => "gra",
        _ => "deg",
    };

    internal static bool TryParseAngles(string value, out AngleUnits angles)
    {
        switch (value)
        {
            case "deg": angles = AngleUnits.Degrees; return true;
            case "rad": angles = AngleUnits.Radians; return true;
            case "gra": angles = AngleUnits.Gradians; return true;
            default: angles = AngleUnits.Degrees; return false;
        }
    }
}
