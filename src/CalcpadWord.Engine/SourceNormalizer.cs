using System.Text;

namespace CalcpadWord.Engine;

/// <summary>
/// Undoes the characters Word's AutoCorrect and AutoFormat substitute while you type
/// (dashes, smart quotes, ½, non-breaking spaces, line breaks), which Calcpad would reject,
/// and turns the operators typed as two keys (>=, ==, &&…) into Calcpad's symbols.
/// </summary>
public static class SourceNormalizer
{
    /// <summary>
    /// Calcpad's own editor shortcuts (as you type, and on paste): == ≡, != ≠, &lt;= ≤, >= ≥,
    /// &amp;&amp; ∧, || ∨, ^^ ⊕, // ÷, %% ⦼, &lt;&lt; ∠, &lt;* ←. Each pair is otherwise an error in Calcpad.
    /// </summary>
    private static char Shortcut(char previous, char c) => (previous, c) switch
    {
        ('=', '=') => '≡',
        ('!', '=') => '≠',
        ('<', '=') => '≤',
        ('>', '=') => '≥',
        ('&', '&') => '∧',
        ('|', '|') => '∨',
        ('^', '^') => '⊕',
        ('/', '/') => '÷',
        ('%', '%') => '⦼',
        ('<', '<') => '∠',
        ('<', '*') => '←',
        ('≡', '=') => '≡', // === and !== (as in JavaScript) mean the same
        ('≠', '=') => '≠',
        _ => '\0',
    };

    private static readonly char[] ShortcutKeys = ['=', '&', '|', '^', '/', '%', '<', '*'];

    // Lines that name a file (a path or URL may contain // or <<).
    private static readonly string[] FileKeywords = ["#include", "#read", "#write", "#append"];

    /// <summary>
    /// Replaces the two-key operators (see <see cref="Shortcut"/>) in one line, leaving comments
    /// alone: Calcpad text between ' or " and the same quote, or the end of the line.
    /// </summary>
    public static string ReplaceOperators(string line)
    {
        if (line.IndexOfAny(ShortcutKeys) < 0)
            return line;

        var trimmed = line.AsSpan().TrimStart();
        foreach (var keyword in FileKeywords)
            if (trimmed.StartsWith(keyword, StringComparison.OrdinalIgnoreCase))
                return line;

        var sb = new StringBuilder(line.Length);
        var comment = '\0';
        var mathStart = 0; // where the current stretch of math began in sb
        foreach (var c in line)
        {
            if (comment != '\0')
            {
                sb.Append(c);
                if (c == comment)
                {
                    comment = '\0';
                    mathStart = sb.Length;
                }
            }
            else if (c is '\'' or '"')
            {
                comment = c;
                sb.Append(c);
            }
            else if (sb.Length > mathStart && Shortcut(sb[^1], c) is var symbol and not '\0')
            {
                sb[^1] = symbol;
            }
            else
            {
                sb.Append(c);
            }
        }
        return sb.ToString();
    }

    /// <summary>One inline expression: line breaks become spaces.</summary>
    public static string Normalize(string? text) => NormalizeLine(text).Trim();

    /// <summary>
    /// A block's script as read from Word: paragraphs (\r), line breaks (\v) and newlines
    /// separate lines. A Windows line ending (\r\n, as in a .cpd file on disk) is one break.
    /// Trailing blank lines are dropped; indentation is kept.
    /// </summary>
    public static string NormalizeBlock(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        var lines = text.Replace("\r\n", "\n").Split(['\r', '\n', '\v', '\f'])
            .Select(line => NormalizeLine(line).TrimEnd())
            .ToList();
        while (lines.Count > 0 && lines[^1].Length == 0)
            lines.RemoveAt(lines.Count - 1);
        while (lines.Count > 0 && lines[0].Length == 0)
            lines.RemoveAt(0);

        return string.Join('\n', lines);
    }

    private static string NormalizeLine(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        var sb = new StringBuilder(text.Length);
        foreach (var c in text)
        {
            switch (c)
            {
                case '–': // en dash: Word turns " - " into " – "
                case '—': // em dash
                case '−': // minus sign
                    sb.Append('-');
                    break;
                case '‘':
                case '’':
                    sb.Append('\'');
                    break;
                case '“':
                case '”':
                    sb.Append('"');
                    break;
                case '½': sb.Append("(1/2)"); break;
                case '¼': sb.Append("(1/4)"); break;
                case '¾': sb.Append("(3/4)"); break;
                case ' ': // non-breaking space
                case ' ': // thin space
                case ' ': // narrow non-breaking space
                case '\r':
                case '\n':
                case '\v':     // Word manual line break
                case '\a':     // end-of-cell marker
                case '\f':
                case '\t':
                    sb.Append(' ');
                    break;
                case '​': // zero-width space
                case '﻿':
                    break;
                default:
                    sb.Append(c);
                    break;
            }
        }
        return ReplaceOperators(sb.ToString());
    }
}
