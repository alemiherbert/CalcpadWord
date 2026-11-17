using System.Text;

namespace CalcpadWord.Engine;

/// <summary>
/// Undoes the characters Word's AutoCorrect and AutoFormat substitute while you type
/// (dashes, smart quotes, ½, non-breaking spaces, line breaks), which Calcpad would reject.
/// </summary>
public static class SourceNormalizer
{
    /// <summary>One inline expression: line breaks become spaces.</summary>
    public static string Normalize(string? text) => NormalizeLine(text).Trim();

    /// <summary>
    /// A block's script as read from Word: paragraphs (\r), line breaks (\v) and newlines
    /// separate lines. Trailing blank lines are dropped; indentation is kept.
    /// </summary>
    public static string NormalizeBlock(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        var lines = text.Split(['\r', '\n', '\v', '\f'])
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
        return sb.ToString();
    }
}
