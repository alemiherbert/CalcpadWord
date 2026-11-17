namespace CalcpadWord.Engine;

/// <summary>
/// Finds blocks typed with fences: a paragraph that is only <c>{=</c>, the script, and a
/// paragraph that is only <c>}</c>. Braces inside the script (e.g. a multi-line
/// <c>$Plot{ … }</c>) are counted, so the closing fence is the first lone <c>}</c> at depth zero.
/// </summary>
public static class BlockFence
{
    public const string Open = "{=";
    public const string Close = "}";
    public const int MaxLines = 5000;

    public static bool IsOpening(string paragraphText) => SourceNormalizer.Normalize(paragraphText) == Open;

    /// <summary>
    /// Index (into <paramref name="following"/>) of the closing fence for an opening fence whose
    /// following paragraphs are given, or -1 if it isn't closed.
    /// </summary>
    public static int FindClose(IEnumerable<string> following)
    {
        var depth = 0;
        var index = 0;
        foreach (var raw in following)
        {
            if (index >= MaxLines)
                break;

            var line = SourceNormalizer.Normalize(raw);
            if (depth == 0 && line == Close)
                return index;
            if (line == Open)
                return -1; // another opening fence first: this one is not closed

            foreach (var c in line)
            {
                if (c == '{') ++depth;
                else if (c == '}' && depth > 0) --depth;
            }
            ++index;
        }
        return -1;
    }
}
