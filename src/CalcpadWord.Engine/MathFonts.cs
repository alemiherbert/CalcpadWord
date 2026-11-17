using System.Text;

namespace CalcpadWord.Engine;

/// <param name="Family">The font family name Word uses (checked against the downloaded files).</param>
/// <param name="Label">How the ribbon shows it.</param>
/// <param name="Note">Where it comes from and its licence.</param>
/// <param name="DownloadUrl">For free fonts: the font file, which the add-in can install for the user.</param>
public sealed record MathFontInfo(string Family, string Label, string Note, string? DownloadUrl = null);

/// <summary>
/// OpenType math fonts offered for equations: these starters (the free ones can be installed for
/// the user), plus any other math font found on the computer. Word can only use an installed one.
/// </summary>
public static class MathFonts
{
    public const string Default = "Cambria Math";

    /// <summary>The note on a math font found on the computer that isn't a starter.</summary>
    public const string FoundNote = "Found on this computer (an OpenType math font).";

    private const string Ctan = "https://mirrors.ctan.org/fonts/";

    public static readonly IReadOnlyList<MathFontInfo> Starters =
    [
        new("Latin Modern Math", "Latin Modern Math", "Free (GUST Font License), from CTAN.", Ctan + "lm-math/opentype/latinmodern-math.otf"),
        new("Cambria Math", "Cambria Math", "Comes with Windows and Office."),
        new("STIX Two Math", "STIX Two Math", "Free (SIL Open Font License), from CTAN.", Ctan + "stix2-otf/STIXTwoMath-Regular.otf"),
        new("Fira Math", "Fira Math", "Free (SIL Open Font License), from CTAN. A sans-serif math font matching Fira Sans.", Ctan + "firamath/FiraMath-Regular.otf"),
        new("Libertinus Math", "Libertinus Math", "Free (SIL Open Font License), from CTAN.", Ctan + "libertinus-fonts/otf/LibertinusMath-Regular.otf"),
        new("XITS Math", "XITS Math", "Free (SIL Open Font License), from CTAN.", Ctan + "xits/XITSMath-Regular.otf"),
        new("TeX Gyre Pagella Math", "TeX Gyre Pagella Math", "Free (GUST Font License), from CTAN.", Ctan + "tex-gyre-math/opentype/texgyrepagella-math.otf"),
        new("Asana Math", "Asana Math", "Free (SIL Open Font License), from CTAN.", Ctan + "Asana-Math/Asana-Math.otf"),
        new("Noto Sans Math", "Noto Sans Math", "Free (SIL Open Font License), from Google Fonts.", "https://github.com/google/fonts/raw/main/ofl/notosansmath/NotoSansMath-Regular.ttf"),
        new("Euler Math", "Euler Math (AMS Euler)", "The free OpenType version of AMS Euler (SIL Open Font License), from CTAN.", Ctan + "euler-math/Euler-Math.otf"),
    ];

    public static MathFontInfo? Find(string family) =>
        Starters.FirstOrDefault(f => f.Family.Equals(family, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// The list to offer, in alphabetical order: every starter (with whether it is installed),
    /// every other math font in <paramref name="installedMath"/>, and <paramref name="current"/>
    /// (the document's font) if it is neither.
    /// </summary>
    public static IReadOnlyList<(MathFontInfo Font, bool Installed)> Choices(
        IReadOnlySet<string> installedMath, Func<string, bool> isInstalled, string current)
    {
        var choices = Starters.Select(f => (Font: f, Installed: installedMath.Contains(f.Family) || isInstalled(f.Family))).ToList();
        foreach (var family in installedMath)
            if (Find(family) is null)
                choices.Add((new MathFontInfo(family, family, FoundNote), true));
        if (!choices.Any(c => c.Font.Family.Equals(current, StringComparison.OrdinalIgnoreCase)))
            choices.Add((new MathFontInfo(current, current, "This document's math font."), isInstalled(current)));
        return [.. choices.OrderBy(c => c.Font.Label, StringComparer.CurrentCultureIgnoreCase)];
    }
}

/// <summary>What an OpenType/TrueType file (or collection) contains: its families, and whether each is a math font.</summary>
public static class FontFile
{
    /// <summary>The fonts in the file as (family, has a MATH table); empty if it isn't a font file.</summary>
    public static IReadOnlyList<(string Family, bool IsMath)> Read(ReadOnlySpan<byte> data)
    {
        var fonts = new List<(string, bool)>();
        if (data.Length < 12)
            return fonts;

        var tag = Encoding.ASCII.GetString(data[..4]);
        if (tag == "ttcf")
        {
            var count = U32(data, 8);
            for (var i = 0; i < count && 12 + 4 * i + 4 <= data.Length; ++i)
                if (ReadFont(data, (int)U32(data, 12 + 4 * i)) is { } font)
                    fonts.Add(font);
        }
        else if (ReadFont(data, 0) is { } font)
        {
            fonts.Add(font);
        }
        return fonts;
    }

    /// <summary>
    /// The math font families in a font file (or collection), reading only the table directories,
    /// and the name table of fonts that have a MATH table: quick enough to look through every
    /// installed font. Empty if it isn't a font file.
    /// </summary>
    public static IReadOnlyList<string> MathFamilies(Stream file)
    {
        var families = new List<string>();
        if (ReadAt(file, 0, 12) is not { } head)
            return families;

        if (Encoding.ASCII.GetString(head, 0, 4) == "ttcf")
        {
            var count = (int)Math.Min(U32(head, 8), 256);
            if (ReadAt(file, 12, 4 * count) is not { } offsets)
                return families;
            for (var i = 0; i < count; ++i)
                if (MathFamily(file, U32(offsets, 4 * i)) is { } family && !families.Contains(family))
                    families.Add(family);
        }
        else if (MathFamily(file, 0) is { } family)
        {
            families.Add(family);
        }
        return families;
    }

    private static string? MathFamily(Stream file, long offset)
    {
        if (ReadAt(file, offset, 12) is not { } header || U32(header, 0) is not (0x00010000 or 0x4F54544F))
            return null;

        var tableCount = U16(header, 4);
        if (ReadAt(file, offset + 12, 16 * tableCount) is not { } records)
            return null;

        long nameOffset = -1, nameLength = 0;
        var hasMath = false;
        for (var i = 0; i < tableCount; ++i)
        {
            var tag = Encoding.ASCII.GetString(records, 16 * i, 4);
            if (tag == "MATH")
                hasMath = true;
            else if (tag == "name")
                (nameOffset, nameLength) = (U32(records, 16 * i + 8), U32(records, 16 * i + 12));
        }
        if (!hasMath || nameOffset < 0 || nameLength is < 6 or > 1 << 20)
            return null;

        return ReadAt(file, nameOffset, (int)nameLength) is { } name && Family(name, 0) is { Length: > 0 } family ? family : null;
    }

    private static byte[]? ReadAt(Stream file, long offset, int length)
    {
        if (offset < 0 || length < 0 || offset + length > file.Length)
            return null;
        var buffer = new byte[length];
        file.Position = offset;
        file.ReadExactly(buffer);
        return buffer;
    }

    private static (string, bool)? ReadFont(ReadOnlySpan<byte> data, int offset)
    {
        if (offset + 12 > data.Length)
            return null;

        var version = U32(data, offset);
        if (version is not (0x00010000 or 0x4F54544F)) // TrueType, or "OTTO" (CFF)
            return null;

        var tableCount = U16(data, offset + 4);
        int nameTable = -1;
        var hasMath = false;
        for (var i = 0; i < tableCount; ++i)
        {
            var record = offset + 12 + 16 * i;
            if (record + 16 > data.Length)
                return null;
            var tag = Encoding.ASCII.GetString(data.Slice(record, 4));
            if (tag == "name") nameTable = (int)U32(data, record + 8);
            else if (tag == "MATH") hasMath = true;
        }
        return nameTable < 0 ? null : (Family(data, nameTable), hasMath);
    }

    /// <summary>
    /// The Windows (platform 3) typographic family (name 16), or the family (name 1); in US
    /// English if the font has it, as Word lists font names.
    /// </summary>
    private static string Family(ReadOnlySpan<byte> data, int table)
    {
        const int UsEnglish = 0x0409;
        if (table + 6 > data.Length)
            return string.Empty;
        var count = U16(data, table + 2);
        var strings = table + U16(data, table + 4);
        string? family = null, typographic = null;
        bool familyIsEnglish = false, typographicIsEnglish = false;
        for (var i = 0; i < count; ++i)
        {
            var record = table + 6 + 12 * i;
            if (record + 12 > data.Length || U16(data, record) != 3)
                continue;
            var id = U16(data, record + 6);
            var english = U16(data, record + 4) == UsEnglish;
            var length = U16(data, record + 8);
            var start = strings + U16(data, record + 10);
            if (id is not (1 or 16) || start + length > data.Length)
                continue;
            var value = Encoding.BigEndianUnicode.GetString(data.Slice(start, length));
            if (id == 16 && (typographic is null || english && !typographicIsEnglish))
                (typographic, typographicIsEnglish) = (value, english);
            else if (id == 1 && (family is null || english && !familyIsEnglish))
                (family, familyIsEnglish) = (value, english);
        }
        return typographic ?? family ?? string.Empty;
    }

    private static int U16(ReadOnlySpan<byte> d, int o) => (d[o] << 8) | d[o + 1];

    private static uint U32(ReadOnlySpan<byte> d, int o) => ((uint)d[o] << 24) | ((uint)d[o + 1] << 16) | ((uint)d[o + 2] << 8) | d[o + 3];
}
