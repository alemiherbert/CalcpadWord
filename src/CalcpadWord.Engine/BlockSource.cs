namespace CalcpadWord.Engine;

/// <summary>Checks on multi-line Calcpad scripts (blocks) and on text that might be one.</summary>
public static class BlockSource
{
    /// <summary>A line ending in one of these (or " _") continues on the next line in Calcpad.</summary>
    private const string ContinuationChars = ";|&@:({[";

    public static IReadOnlyList<string> Lines(string source) =>
        source.Length == 0 ? [] : source.Split('\n');

    /// <summary>
    /// Why a block can't run as part of the document, or null. A block must close every
    /// #if, loop and #def it opens, so it can't change how the rest of the document runs.
    /// </summary>
    public static string? Problem(IReadOnlyList<string> lines)
    {
        if (lines.All(l => l.Trim().Length == 0))
            return "The block is empty. Type Calcpad code in it, one statement per line.";

        int ifs = 0, loops = 0, defs = 0;
        foreach (var raw in lines)
        {
            var line = raw.Trim();
            if (!line.StartsWith('#'))
                continue;

            var lower = line.ToLowerInvariant();
            if (IsKeyword(lower, "#pause") || IsKeyword(lower, "#input"))
                return "#pause and #input are not supported in Word.";
            if (IsKeyword(lower, "#end if"))
            {
                if (--ifs < 0) return "\"#end if\" without a matching \"#if\".";
            }
            else if (IsKeyword(lower, "#end def"))
            {
                if (--defs < 0) return "\"#end def\" without a matching \"#def\".";
            }
            else if (IsKeyword(lower, "#if"))
                ++ifs;
            else if (IsKeyword(lower, "#loop"))
            {
                if (--loops < 0) return "\"#loop\" without a matching \"#for\", \"#while\" or \"#repeat\".";
            }
            else if (IsKeyword(lower, "#for") || IsKeyword(lower, "#while") || IsKeyword(lower, "#repeat"))
                ++loops;
            else if (IsKeyword(lower, "#def") && !line.Contains('='))
                ++defs; // multi-line macro
        }
        if (ifs > 0) return "\"#if\" is not closed. Add \"#end if\" inside the block.";
        if (loops > 0) return "A loop is not closed. Add \"#loop\" inside the block.";
        if (defs > 0) return "A macro is not closed. Add \"#end def\" inside the block.";

        var last = lines.Last(l => l.Trim().Length > 0).Trim();
        if (!IsComment(last) && EndsWithContinuation(last))
            return "The last line is incomplete: Calcpad would join it with whatever follows the block.";

        return null;
    }

    /// <summary>
    /// True when the lines read like a Calcpad script rather than a list of expressions:
    /// some line is a comment, a heading, a keyword or a plot.
    /// </summary>
    public static bool LooksLikeScript(IEnumerable<string> lines)
    {
        foreach (var raw in lines)
        {
            var line = SourceNormalizer.Normalize(raw);
            if (line.Length == 0)
                continue;
            if (line[0] is '\'' or '"')
                return true;
            if (line.StartsWith("$plot", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("$map", StringComparison.OrdinalIgnoreCase))
                return true;
            if (line[0] == '#')
            {
                // "#val M" is an inline expression with a switch; "#if x > 0" or a lone "#round 3" is a keyword line.
                try
                {
                    if (InlineDirective.Parse(line).Expression.Length == 0)
                        return true;
                }
                catch (FormatException)
                {
                    return true;
                }
            }
        }
        return false;
    }

    internal static bool EndsWithContinuation(string line) =>
        line.EndsWith(" _", StringComparison.Ordinal) ||
        line.Length > 0 && ContinuationChars.Contains(line[^1]);

    internal static bool IsComment(string line) => line.Length > 0 && line[0] is '\'' or '"';

    private static bool IsKeyword(string lowerLine, string keyword) =>
        lowerLine.StartsWith(keyword, StringComparison.Ordinal) &&
        (lowerLine.Length == keyword.Length || !char.IsLetterOrDigit(lowerLine[keyword.Length]));
}
