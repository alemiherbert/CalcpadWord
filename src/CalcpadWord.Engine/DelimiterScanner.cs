namespace CalcpadWord.Engine;

public readonly record struct DelimitedExpression(int Start, int Length, string Source);

/// <summary>
/// Finds inline expressions written as <c>{= ... }</c>. Braces inside are matched, so
/// solver blocks such as <c>{= $Root{x^2 - 2 @ x = 0 : 2} }</c> work. An expression must
/// close within its paragraph; otherwise it is left alone.
/// </summary>
public static class DelimiterScanner
{
    public const string Open = "{=";

    /// <summary>Length of the delimited expression starting at <paramref name="start"/>, or -1.</summary>
    public static int MatchLength(string text, int start)
    {
        if (start < 0 || !text.AsSpan(start).StartsWith(Open, StringComparison.Ordinal))
            return -1;

        var depth = 1;
        for (int i = start + Open.Length; i < text.Length; ++i)
        {
            switch (text[i])
            {
                case '\r' or '\n' or '\a' or '\f':
                    return -1;
                case '{':
                    ++depth;
                    break;
                case '}':
                    if (--depth == 0)
                        return i - start + 1;
                    break;
            }
        }
        return -1;
    }

    /// <summary>The source between <c>{=</c> and the matching <c>}</c>.</summary>
    public static string Inner(string delimited) => delimited[Open.Length..^1].Trim();

    public static IReadOnlyList<DelimitedExpression> FindAll(string text)
    {
        var found = new List<DelimitedExpression>();
        var i = 0;
        while ((i = text.IndexOf(Open, i, StringComparison.Ordinal)) >= 0)
        {
            var length = MatchLength(text, i);
            if (length < 0)
            {
                i += Open.Length;
                continue;
            }
            found.Add(new DelimitedExpression(i, length, Inner(text.Substring(i, length))));
            i += length;
        }
        return found;
    }
}
