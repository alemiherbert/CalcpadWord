using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace CalcpadWord.Engine;

/// <summary>Reading Calcpad's HTML output.</summary>
internal static partial class CalcpadHtml
{
    [GeneratedRegex(@"id=""eq-(\d+)""")]
    public static partial Regex EquationId();

    // " on line [<a href="#0" data-text="12">12</a>]": script line numbers mean nothing in Word.
    [GeneratedRegex(@" on line \[<a [^>]*>\d+</a>\]")]
    public static partial Regex LineReference();

    [GeneratedRegex(@" on line \[\d+\]:\s*(?<message>.*)$", RegexOptions.Singleline)]
    private static partial Regex ErrorMessage();

    [GeneratedRegex(@"\s+")]
    private static partial Regex Whitespace();

    public static HtmlNode Parse(string html)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(html);
        return doc.DocumentNode;
    }

    /// <summary>Calcpad error elements, not counting the "!!!" warnings inside equations.</summary>
    public static IReadOnlyList<HtmlNode> Errors(HtmlNode root) =>
        root.Descendants()
            .Where(n => n.HasClass("err") && !n.Ancestors().Any(a => a.HasClass("eq")))
            .ToList();

    /// <summary>The message of a Calcpad error, without the "Error in "…" on line [n]:" preamble.</summary>
    public static string ErrorText(HtmlNode error)
    {
        var text = Text(error);
        var m = ErrorMessage().Match(text);
        return m.Success ? m.Groups["message"].Value.Trim() : text;
    }

    public static IReadOnlyList<HtmlNode> Equations(HtmlNode root) =>
        root.Descendants("span").Where(n => n.HasClass("eq")).ToList();

    public static int EquationIndex(HtmlNode equation) =>
        int.TryParse(equation.GetAttributeValue("id", "").Replace("eq-", ""), out var i) ? i : -1;

    /// <summary>Readable text: powers as ^, subscripts as _, fractions as (a/b).</summary>
    public static string Text(HtmlNode node)
    {
        var sb = new StringBuilder();
        Append(node, sb);
        return Whitespace().Replace(sb.ToString(), " ").Trim();

        static void Append(HtmlNode n, StringBuilder sb)
        {
            if (n.NodeType == HtmlNodeType.Text)
            {
                sb.Append(HtmlEntity.DeEntitize(n.InnerText));
                return;
            }
            if (n.NodeType == HtmlNodeType.Comment)
                return;

            if (n.Name == "sup") sb.Append('^');
            else if (n.Name == "sub") sb.Append('_');
            else if (n.HasClass("dvl")) { sb.Append('/'); return; }
            else if (n.HasClass("dvc")) sb.Append('(');
            else if (n.Name is "p" or "h3" or "br" or "div" or "tr") sb.Append(' ');

            foreach (var child in n.ChildNodes)
                Append(child, sb);

            if (n.HasClass("dvc")) sb.Append(')');
        }
    }
}
