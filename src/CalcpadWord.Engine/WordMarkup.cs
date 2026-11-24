using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace CalcpadWord.Engine;

public enum TextLook
{
    /// <summary>Calcpad source being edited.</summary>
    Source,
    /// <summary>Calcpad source that failed to evaluate.</summary>
    Error,
    /// <summary>Source of a <c>#hide</c> expression: Word hidden text.</summary>
    Hidden,
}

/// <summary>
/// Builds the WordprocessingML that goes into a content control via <c>Range.InsertXML</c>.
/// Everything is inserted this way: setting <c>Range.Text</c> on a control that holds an
/// equation would put the text inside the math zone.
/// </summary>
public static partial class WordMarkup
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string AssignmentRun = "<m:r><m:t>=</m:t></m:r>";
    // Consolas, and excluded from spelling and grammar checks (Calcpad code is not prose).
    // Run properties must stay in schema order: rFonts, noProof, vanish, color, u, shd.
    private const string SourceFont = "<w:rFonts w:ascii=\"Consolas\" w:hAnsi=\"Consolas\" w:cs=\"Consolas\"/><w:noProof/>";

    [GeneratedRegex(@"<w:sz\s+w:val=""\d+""\s*/>")]
    private static partial Regex FontSize();

    [GeneratedRegex(@"<w:color\s+w:val=""[0-9A-Fa-f]{6}""\s*/>")]
    private static partial Regex Colour();

    [GeneratedRegex(@"<w:rPr>\s*</w:rPr>")]
    private static partial Regex EmptyRunProperties();

    // A math run, <m:r> (not <m:rPr>); runs don't nest.
    [GeneratedRegex(@"<m:r(?:\s[^>]*)?>.*?</m:r>", RegexOptions.Singleline)]
    private static partial Regex MathRun();

    [GeneratedRegex(@"<m:rPr>\s*<m:nor\s*/>")]
    private static partial Regex NormalText();

    // Calcpad writes the × of 5.7×10⁻⁹ as a small, bold, superscript normal-text run (raised in Word).
    [GeneratedRegex(@"<m:r>\s*<m:rPr>\s*<m:nor\s*/>\s*</m:rPr>\s*<w:rPr>(?:(?!</w:rPr>).)*?<w:vertAlign\s+w:val=""superscript""\s*/>(?:(?!</w:rPr>).)*?</w:rPr>\s*<m:t>×</m:t>\s*</m:r>", RegexOptions.Singleline)]
    private static partial Regex RaisedTimes();

    // ...and every minus sign (a negative exponent, -a, exp(-1)) as a hyphen in normal text, which Word
    // draws as a hyphen in the paragraph's font. A hyphen in a math run is drawn as a minus already.
    [GeneratedRegex(@"<m:r>\s*<m:rPr>\s*<m:nor\s*/>\s*</m:rPr>\s*(?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>\s*)?<m:t>-</m:t>\s*</m:r>", RegexOptions.Singleline)]
    private static partial Regex HyphenMinus();

    // Function names (sin, max, exp…): Calcpad writes them bold; Word math and TeX set them upright and regular.
    [GeneratedRegex(@"(<m:r>\s*<m:rPr>\s*<m:nor\s*/>\s*</m:rPr>\s*<w:rPr>(?:(?!</w:rPr>).)*?)<w:b(?:\s+w:val=""(?:true|1|on)"")?\s*/>", RegexOptions.Singleline)]
    private static partial Regex BoldFunction();

    // The subscript of a name (M_Ed, A_s,req): a script's m:sub, after its base (m:e); not the limits of a sum.
    [GeneratedRegex(@"(?<=</m:e>\s*)<m:sub>(?:(?!</m:sub>).)*</m:sub>", RegexOptions.Singleline)]
    private static partial Regex Subscript();

    // A math run with no properties: in a subscript, the name's letters (a variable there, an index, has Calcpad's colour).
    [GeneratedRegex(@"<m:r>(?=\s*<m:t>[^<]*\p{L})")]
    private static partial Regex PlainRun();

    private const string TimesRun = "<m:r><m:t>×</m:t></m:r>";
    private const string MinusRun = "<m:r><m:t>−</m:t></m:r>";
    private const string UprightRun = "<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr>";

    /// <summary>
    /// Calcpad raises the × of scientific notation and writes minus signs as hyphens in normal text:
    /// both become ordinary math characters (5.7 × 10⁻⁹, −a, exp(−1)).
    /// Calcpad hard-codes 10–11 pt for units; dropping sizes lets equations follow the paragraph.
    /// Calcpad writes Cambria Math on units and other plain text in equations, and no font on the
    /// rest, which Word then shows in Cambria Math whatever the document's math font: with another
    /// math font, every run gets that font. Normal text without a font ("if", "else:") would be drawn
    /// in the paragraph's font, so it always gets the math font.
    /// Names' subscripts are upright, as in Calcpad's own output and the Eurocodes (M_Ed, A_s,req), and
    /// function names regular, as in Word math and TeX (Calcpad writes them bold).
    /// The plain style drops Calcpad's blue variable colour.
    /// </summary>
    public static string ApplyStyle(string omml, RenderStyle style, string mathFont = MathFonts.Default)
    {
        // Scientific notation as Word math writes it: 5.7 × 10^(−9), on the line, with a real minus.
        var s = RaisedTimes().Replace(omml, TimesRun);
        s = HyphenMinus().Replace(s, MinusRun);
        s = BoldFunction().Replace(s, "$1");
        s = Subscript().Replace(s, sub => PlainRun().Replace(sub.Value, UprightRun));
        s = FontSize().Replace(s, string.Empty);
        s = mathFont == MathFonts.Default
            ? WithMathFont(s, mathFont, normalTextOnly: true)
            : WithMathFont(s.Replace("\"Cambria Math\"", "\"" + SecurityElement.Escape(mathFont) + "\"", StringComparison.Ordinal), mathFont);
        if (style == RenderStyle.PlainMath)
            s = Colour().Replace(s, string.Empty);

        return EmptyRunProperties().Replace(s, string.Empty);
    }

    /// <summary>
    /// Gives each math run without a font the chosen one (w:rFonts comes first in w:rPr, after a w:rStyle).
    /// With Cambria Math, Word's default, only normal text needs it: math runs are drawn in the math font anyway.
    /// </summary>
    private static string WithMathFont(string omml, string mathFont, bool normalTextOnly = false)
    {
        var font = SecurityElement.Escape(mathFont);
        var rFonts = $"<w:rFonts w:ascii=\"{font}\" w:hAnsi=\"{font}\"/>";
        return MathRun().Replace(omml, m =>
        {
            var run = m.Value;
            if (run.Contains("<w:rFonts", StringComparison.Ordinal) || normalTextOnly && !NormalText().IsMatch(run))
                return run;

            var props = run.IndexOf("<w:rPr>", StringComparison.Ordinal);
            if (props >= 0)
            {
                var at = props + "<w:rPr>".Length;
                var style = Regex.Match(run[at..], @"^<w:rStyle\b[^>]*/>");
                return run.Insert(at + style.Length, rFonts);
            }
            if (run.Contains("<w:rPr/>", StringComparison.Ordinal))
                return run.Replace("<w:rPr/>", $"<w:rPr>{rFonts}</w:rPr>", StringComparison.Ordinal);

            var text = run.IndexOf("<m:t", StringComparison.Ordinal); // w:rPr goes after m:rPr, before the text
            return text < 0 ? run : run.Insert(text, $"<w:rPr>{rFonts}</w:rPr>");
        });
    }

    /// <summary>The part after the last "=" (the result), or the whole equation if there is none.</summary>
    public static string ResultOnly(string omml)
    {
        var i = omml.LastIndexOf(AssignmentRun, StringComparison.Ordinal);
        if (i < 0)
            return omml;

        var tail = omml[(i + AssignmentRun.Length)..];
        return IsWellFormedMath(tail) ? tail : omml;
    }

    public static string TextResultOnly(string text)
    {
        var i = text.LastIndexOf(" = ", StringComparison.Ordinal);
        return i < 0 ? text : text[(i + 3)..];
    }

    public static string MathParagraph(string omml) => $"<w:p><m:oMath>{omml}</m:oMath></w:p>";

    public static string TextParagraph(string text, TextLook look)
    {
        var props = look switch
        {
            TextLook.Error => $"{SourceFont}<w:color w:val=\"C00000\"/><w:u w:val=\"wave\" w:color=\"FF0000\"/>",
            TextLook.Hidden => $"{SourceFont}<w:vanish/><w:color w:val=\"7F7F7F\"/>",
            _ => $"{SourceFont}<w:color w:val=\"1F4E79\"/><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"EEF3F8\"/>",
        };
        return $"<w:p><w:r><w:rPr>{props}</w:rPr><w:t xml:space=\"preserve\">{SecurityElement.Escape(text)}</w:t></w:r></w:p>";
    }

    /// <summary>A block's source: one paragraph per line (an empty block gets one empty paragraph to type in).</summary>
    public static string SourceParagraphs(string source, TextLook look)
    {
        var lines = source.Length == 0 ? [string.Empty] : source.Split('\n');
        return string.Concat(lines.Select(line => TextParagraph(line, look)));
    }

    /// <summary>
    /// A converted script (<see cref="ScriptConverter"/>) as body content. Expressions and blocks are
    /// content controls already tagged with their source, so the next recalculation draws them.
    /// Headings carry only an outline level: the add-in then applies Word's built-in Heading style,
    /// which works in every language and keeps the document's own heading look.
    /// </summary>
    public static string ConvertedContent(IReadOnlyList<DocumentPart> parts)
    {
        var sb = new StringBuilder();
        foreach (var part in parts)
        {
            switch (part)
            {
                case RulePart:
                    sb.Append("<w:p><w:pPr><w:pBdr><w:bottom w:val=\"single\" w:sz=\"6\" w:space=\"1\" w:color=\"auto\"/></w:pBdr></w:pPr></w:p>");
                    break;

                case BlockPart block:
                    var blockTag = new ControlTag(ControlState.Result, string.Empty, block.Source, ControlKind.Block).ToString();
                    sb.Append("<w:sdt><w:sdtPr>")
                      .Append($"<w:alias w:val=\"{Attribute(Clip("Calcpad block: " + block.Source.Split('\n')[0]))}\"/>")
                      .Append($"<w:tag w:val=\"{Attribute(blockTag)}\"/><w:id w:val=\"{NextId()}\"/>")
                      .Append("</w:sdtPr><w:sdtContent>")
                      .Append(SourceParagraphs(block.Source, TextLook.Source))
                      .Append("</w:sdtContent></w:sdt>");
                    break;

                case ParagraphPart paragraph:
                    sb.Append("<w:p>");
                    if (paragraph.Hidden)
                        sb.Append("<w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr>");
                    else if (paragraph.HeadingLevel > 0)
                        sb.Append($"<w:pPr><w:outlineLvl w:val=\"{paragraph.HeadingLevel - 1}\"/></w:pPr>");

                    foreach (var run in paragraph.Runs)
                        sb.Append(Run(run, paragraph.Hidden));
                    sb.Append("</w:p>");
                    break;
            }
        }
        return sb.ToString();

        static string Run(RunPart run, bool hidden)
        {
            switch (run)
            {
                case BreakRun:
                    return "<w:r><w:br/></w:r>";

                case InputRun input:
                    // A plain-text form field in the theme's accent colour; the tag names the variable.
                    var field = new InputField(input.Name, input.Unit);
                    var look = $"<w:color w:val=\"2E63C4\" w:themeColor=\"accent1\"/>{(hidden ? "<w:vanish/>" : "")}";
                    return "<w:sdt><w:sdtPr>" +
                           $"<w:rPr>{look}</w:rPr>" +
                           $"<w:alias w:val=\"{Attribute(Clip("Input: " + input.Name + (input.Unit.Length > 0 ? $" ({input.Unit})" : "")))}\"/>" +
                           $"<w:tag w:val=\"{Attribute(field.ToString())}\"/><w:id w:val=\"{NextId()}\"/><w:text/>" +
                           "</w:sdtPr><w:sdtContent>" +
                           $"<w:r><w:rPr>{look}</w:rPr><w:t xml:space=\"preserve\">{SecurityElement.Escape(input.Value.Length == 0 ? "?" : input.Value)}</w:t></w:r>" +
                           "</w:sdtContent></w:sdt>";

                case ExpressionRun expression:
                    var tag = new ControlTag(ControlState.Result, string.Empty, expression.Source).ToString();
                    return "<w:sdt><w:sdtPr>" +
                           $"<w:alias w:val=\"{Attribute(Clip("Calcpad: " + expression.Source))}\"/>" +
                           $"<w:tag w:val=\"{Attribute(tag)}\"/><w:id w:val=\"{NextId()}\"/>" +
                           "</w:sdtPr><w:sdtContent>" +
                           $"<w:r><w:rPr>{SourceFont}{(hidden ? "<w:vanish/>" : "")}</w:rPr><w:t xml:space=\"preserve\">{SecurityElement.Escape(expression.Source)}</w:t></w:r>" +
                           "</w:sdtContent></w:sdt>";

                default:
                    var text = (TextRun)run;
                    var props = new StringBuilder();
                    if (text.Format.HasFlag(TextFormat.Bold)) props.Append("<w:b/>");
                    if (text.Format.HasFlag(TextFormat.Italic)) props.Append("<w:i/>");
                    if (hidden) props.Append("<w:vanish/>");
                    if (text.Format.HasFlag(TextFormat.Underline)) props.Append("<w:u w:val=\"single\"/>");
                    if (text.Format.HasFlag(TextFormat.Subscript)) props.Append("<w:vertAlign w:val=\"subscript\"/>");
                    else if (text.Format.HasFlag(TextFormat.Superscript)) props.Append("<w:vertAlign w:val=\"superscript\"/>");
                    var rPr = props.Length == 0 ? "" : $"<w:rPr>{props}</w:rPr>";
                    return $"<w:r>{rPr}<w:t xml:space=\"preserve\">{SecurityElement.Escape(text.Text)}</w:t></w:r>";
            }
        }
    }

    private static int _nextId = Random.Shared.Next(1_000_000, 100_000_000);

    private static int NextId() => Interlocked.Increment(ref _nextId);

    private static string Clip(string s) => s.Length <= 250 ? s : s[..249] + "…";

    /// <summary>An XML attribute value: line breaks must be character references, or XML turns them into spaces.</summary>
    private static string Attribute(string value) =>
        SecurityElement.Escape(value)!.Replace("\r", "&#13;").Replace("\n", "&#10;").Replace("\t", "&#9;");

    /// <summary>Wraps body content in a Flat OPC package, the format <c>Range.InsertXML</c> accepts.</summary>
    public static string Package(string bodyXml)
    {
        var sb = new StringBuilder(bodyXml.Length + 1200);
        sb.Append("""
            <?xml version="1.0" standalone="yes"?>
            <?mso-application progid="Word.Document"?>
            <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
            <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
            </pkg:xmlData></pkg:part>
            <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>
            """);
        sb.Append($"<w:document xmlns:w=\"{W}\" xmlns:m=\"{M}\"><w:body>");
        sb.Append(bodyXml);
        sb.Append("</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>");
        return sb.ToString();
    }

    internal static bool IsWellFormedMath(string fragment)
    {
        try
        {
            XElement.Parse($"<x xmlns:w=\"{W}\" xmlns:m=\"{M}\">{fragment}</x>");
            return true;
        }
        catch (System.Xml.XmlException)
        {
            return false;
        }
    }
}
